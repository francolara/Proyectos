#!/usr/bin/env bash
# Instala de forma transaccional el despliegue restringido en el VPS.
set -Eeuo pipefail

readonly DEPLOY_USER="fralse-deploy"
readonly INSTALL_DIR="/usr/local/sbin"
readonly COMPOSE_DIR="/srv/fralse/compose"
readonly SUDOERS_FILE="/etc/sudoers.d/fralse-deploy"
readonly BACKUP_ID="$(date -u '+%Y%m%dT%H%M%SZ').$$"

usage() {
    echo "Uso: sudo bash install-deploy.sh --public-key-file <ruta>" >&2
    exit 64
}

PUBLIC_KEY_FILE=""
while [[ $# -gt 0 ]]; do
    case "$1" in
        --public-key-file) [[ $# -ge 2 ]] || usage; PUBLIC_KEY_FILE="$2"; shift 2 ;;
        *) usage ;;
    esac
done

[[ $EUID -eq 0 ]] || { echo "Este instalador debe ejecutarse como root." >&2; exit 1; }
[[ -n "$PUBLIC_KEY_FILE" && -r "$PUBLIC_KEY_FILE" ]] || {
    echo "No se puede leer la clave pública indicada." >&2
    exit 1
}

SOURCE_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd -P)"
for source_file in fralse-deploy fralse-deploy-ssh ghcr.override.yml; do
    [[ -f "$SOURCE_DIR/$source_file" && -r "$SOURCE_DIR/$source_file" ]] || {
        echo "Falta o no se puede leer $source_file." >&2
        exit 1
    }
done
bash -n "$SOURCE_DIR/fralse-deploy"
bash -n "$SOURCE_DIR/fralse-deploy-ssh"
[[ -d "$COMPOSE_DIR" ]] || { echo "No existe $COMPOSE_DIR." >&2; exit 1; }

PUBLIC_KEY="$(tr -d '\r\n' < "$PUBLIC_KEY_FILE")"
[[ "$PUBLIC_KEY" =~ ^(ssh-ed25519|ssh-rsa|ecdsa-sha2-nistp(256|384|521))[[:space:]][A-Za-z0-9+/=]+([[:space:]].*)?$ ]] || {
    echo "La entrada no tiene un formato de clave pública SSH permitido." >&2
    exit 1
}

STAGE_DIR="$(mktemp -d /root/fralse-deploy-install.XXXXXX)"
RESTORE_DIR="$STAGE_DIR/restore"
mkdir -p "$RESTORE_DIR"
INSTALL_STARTED=0
ACCOUNT_CREATED=0
SSH_DIR_EXISTED=0
declare -A TARGET_EXISTED=()

cleanup_stage() {
    rm -rf -- "$STAGE_DIR"
}
restore_target() {
    local label="$1" destination="$2"
    if [[ "${TARGET_EXISTED[$label]:-0}" == 1 ]]; then
        rm -f -- "$destination"
        cp -a -- "$RESTORE_DIR/$label" "$destination"
    else
        rm -f -- "$destination"
    fi
}
restore_installation() {
    set +e
    restore_target deploy_script "$INSTALL_DIR/fralse-deploy"
    restore_target ssh_wrapper "$INSTALL_DIR/fralse-deploy-ssh"
    restore_target ghcr_override "$COMPOSE_DIR/ghcr.override.yml"
    restore_target sudoers "$SUDOERS_FILE"
    restore_target authorized_keys "$AUTHORIZED_KEYS"
    if (( SSH_DIR_EXISTED == 0 )); then
        rmdir -- "$SSH_DIR" 2>/dev/null || true
    fi
    if (( ACCOUNT_CREATED == 1 )); then
        userdel --remove "$DEPLOY_USER" 2>/dev/null || true
    fi
}
on_error() {
    local status="$1"
    trap - ERR
    if (( INSTALL_STARTED == 1 )); then
        echo "La instalación falló; restaurando el estado anterior." >&2
        restore_installation
    elif (( ACCOUNT_CREATED == 1 )); then
        userdel --remove "$DEPLOY_USER" 2>/dev/null || true
    fi
    cleanup_stage
    exit "$status"
}
trap 'on_error $?' ERR
trap cleanup_stage EXIT

# Valida y prepara todo el contenido antes de modificar el sistema.
install -o root -g root -m 0750 "$SOURCE_DIR/fralse-deploy" "$STAGE_DIR/fralse-deploy"
install -o root -g root -m 0755 "$SOURCE_DIR/fralse-deploy-ssh" "$STAGE_DIR/fralse-deploy-ssh"
install -o root -g root -m 0640 "$SOURCE_DIR/ghcr.override.yml" "$STAGE_DIR/ghcr.override.yml"

printf '%s ALL=(root) NOPASSWD: %s/fralse-deploy\n' "$DEPLOY_USER" "$INSTALL_DIR" > "$STAGE_DIR/sudoers"
chown root:root "$STAGE_DIR/sudoers"
chmod 0440 "$STAGE_DIR/sudoers"
visudo -cf "$STAGE_DIR/sudoers"

if ! id "$DEPLOY_USER" >/dev/null 2>&1; then
    useradd --create-home --user-group --shell /bin/bash --password '!' "$DEPLOY_USER"
    ACCOUNT_CREATED=1
fi
[[ "$DEPLOY_USER" != root ]] || { echo "La cuenta de despliegue nunca puede ser root." >&2; exit 1; }
for group_name in $(id -nG "$DEPLOY_USER"); do
    case "$group_name" in
        docker|sudo|wheel|adm)
            echo "La cuenta $DEPLOY_USER pertenece al grupo privilegiado $group_name." >&2
            exit 1
            ;;
    esac
done
SHADOW_ENTRY="$(getent shadow "$DEPLOY_USER")"
SHADOW_PASSWORD="${SHADOW_ENTRY#*:}"
SHADOW_PASSWORD="${SHADOW_PASSWORD%%:*}"
[[ "$SHADOW_PASSWORD" == '!'* || "$SHADOW_PASSWORD" == '*'* ]] || {
    echo "La cuenta $DEPLOY_USER tiene una contraseña utilizable; bloquéela antes de instalar." >&2
    exit 1
}

USER_HOME="$(getent passwd "$DEPLOY_USER" | cut -d: -f6)"
USER_GROUP="$(id -gn "$DEPLOY_USER")"
[[ -n "$USER_HOME" && -d "$USER_HOME" ]] || { echo "La cuenta de despliegue no tiene directorio personal válido." >&2; exit 1; }
SSH_DIR="$USER_HOME/.ssh"
AUTHORIZED_KEYS="$SSH_DIR/authorized_keys"
[[ -d "$SSH_DIR" ]] && SSH_DIR_EXISTED=1

KEY_TYPE="${PUBLIC_KEY%% *}"
KEY_REMAINDER="${PUBLIC_KEY#* }"
KEY_BLOB="${KEY_REMAINDER%% *}"
KEY_MATERIAL="$KEY_TYPE $KEY_BLOB"
FORCED_KEY="command=\"${INSTALL_DIR}/fralse-deploy-ssh\",no-port-forwarding,no-agent-forwarding,no-X11-forwarding,no-pty ${PUBLIC_KEY}"
if [[ -f "$AUTHORIZED_KEYS" ]]; then
    cp -a -- "$AUTHORIZED_KEYS" "$STAGE_DIR/authorized_keys"
else
    : > "$STAGE_DIR/authorized_keys"
fi
if ! grep -Fqx -- "$FORCED_KEY" "$STAGE_DIR/authorized_keys"; then
    if grep -Fq -- "$KEY_MATERIAL" "$STAGE_DIR/authorized_keys"; then
        echo "La misma clave ya existe sin la entrada restringida esperada; revísela antes de continuar." >&2
        exit 1
    fi
    printf '%s\n' "$FORCED_KEY" >> "$STAGE_DIR/authorized_keys"
fi
chown "$DEPLOY_USER:$USER_GROUP" "$STAGE_DIR/authorized_keys"
chmod 0600 "$STAGE_DIR/authorized_keys"

backup_target() {
    local label="$1" destination="$2"
    if [[ -e "$destination" ]]; then
        TARGET_EXISTED["$label"]=1
        cp -a -- "$destination" "$RESTORE_DIR/$label"
        cp -a -- "$destination" "${destination}.bak.$BACKUP_ID"
    else
        TARGET_EXISTED["$label"]=0
    fi
}
atomic_install() {
    local source_file="$1" destination="$2" owner="$3" group="$4" mode="$5" temporary
    temporary="$(mktemp "${destination}.tmp.XXXXXX")"
    install -o "$owner" -g "$group" -m "$mode" "$source_file" "$temporary"
    mv -f -- "$temporary" "$destination"
}

# Desde este punto cada modificación tiene respaldo y el trap restaura el estado previo.
backup_target deploy_script "$INSTALL_DIR/fralse-deploy"
backup_target ssh_wrapper "$INSTALL_DIR/fralse-deploy-ssh"
backup_target ghcr_override "$COMPOSE_DIR/ghcr.override.yml"
backup_target sudoers "$SUDOERS_FILE"
backup_target authorized_keys "$AUTHORIZED_KEYS"
INSTALL_STARTED=1

install -d -o "$DEPLOY_USER" -g "$USER_GROUP" -m 0700 "$SSH_DIR"
atomic_install "$STAGE_DIR/fralse-deploy" "$INSTALL_DIR/fralse-deploy" root root 0750
atomic_install "$STAGE_DIR/fralse-deploy-ssh" "$INSTALL_DIR/fralse-deploy-ssh" root root 0755
atomic_install "$STAGE_DIR/ghcr.override.yml" "$COMPOSE_DIR/ghcr.override.yml" root root 0640
atomic_install "$STAGE_DIR/sudoers" "$SUDOERS_FILE" root root 0440
visudo -cf "$SUDOERS_FILE"
atomic_install "$STAGE_DIR/authorized_keys" "$AUTHORIZED_KEYS" "$DEPLOY_USER" "$USER_GROUP" 0600
chown "$DEPLOY_USER:$USER_GROUP" "$SSH_DIR"
chmod 0700 "$SSH_DIR"

INSTALL_STARTED=0
echo "Instalación completada para la cuenta exclusiva $DEPLOY_USER. Respaldos: *.bak.$BACKUP_ID"
