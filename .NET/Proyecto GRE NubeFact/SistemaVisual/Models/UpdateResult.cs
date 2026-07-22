namespace SistemaVisual.Models
{
    public sealed class UpdateResult
    {
        public bool Actualizado { get; set; }
        public bool SinCambios { get; set; }
        public bool SinConexion { get; set; }
        public string Mensaje { get; set; }
        public bool ModoInstalacion { get; set; }
        public int Nuevos { get; set; }
        public int Actualizados { get; set; }
        public int SinCambio { get; set; }
        public int Omitidos { get; set; }
    }
}
