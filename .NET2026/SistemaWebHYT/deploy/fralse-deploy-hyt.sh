#!/usr/bin/env bash
# Despliegue de HYT invocado por el comando restringido de FRALSE.
set -Eeuo pipefail
umask 077

tag="${1:-}"
[[ $# -eq 1 && ( "$tag" == main || "$tag" =~ ^[0-9a-f]{40}$ ) ]] || exit 64
readonly image="ghcr.io/francolara/hyt:$tag"
readonly parameters="/srv/fralse/compose/hyt.compose.env"
readonly service="hyt-web"
readonly -a compose=(docker compose --env-file "$parameters" -f /srv/fralse/compose/compose.hyt.yml)

exec 9>/var/lock/fralse-deploy.lock
flock -n 9 || { echo "Ya hay un despliegue en curso." >&2; exit 1; }

[[ "$(grep -c '^HYT_IMAGE=' "$parameters")" == 1 ]]
container="$("${compose[@]}" ps -q "$service")"
[[ -n "$container" ]] || { echo "HYT requiere un contenedor previo para permitir reversión." >&2; exit 1; }
previous="$(docker inspect --format '{{.Image}}' "$container")"
[[ "$previous" =~ ^sha256:[0-9a-f]{64}$ ]]

health_check() {
    local attempt
    for ((attempt = 1; attempt <= 30; attempt++)); do
        if [[ "$(curl -fsS --max-time 5 -H 'Host: hyt.fralsetech.com'             http://127.0.0.1:5186/health/live 2>/dev/null)" == ok ]]; then
            return 0
        fi
        sleep 2
    done
    return 1
}

HYT_IMAGE="$image" "${compose[@]}" config -q
HYT_IMAGE="$image" "${compose[@]}" pull "$service"

next="$(mktemp /srv/fralse/compose/.hyt.compose.XXXXXXXX)"
trap 'rm -f -- "$next"' EXIT
sed "s|^HYT_IMAGE=.*|HYT_IMAGE=$image|" "$parameters" > "$next"
chmod --reference="$parameters" "$next"
chown --reference="$parameters" "$next"

if HYT_IMAGE="$image" "${compose[@]}" up -d --no-deps --pull never "$service"     && health_check && mv -T -- "$next" "$parameters"; then
    printf 'Despliegue HYT completado con etiqueta %s.
' "$tag"
    exit 0
fi

echo "Falló el despliegue de HYT. Restaurando la imagen anterior." >&2
if HYT_IMAGE="$previous" "${compose[@]}" up -d --no-deps --pull never "$service" && health_check; then
    echo "HYT volvió a la imagen anterior." >&2
else
    echo "La reversión de HYT requiere revisión en Contabo." >&2
fi
exit 1
