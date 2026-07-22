namespace SistemaVisual.Models
{
    public sealed class UpdateProgress
    {
        public string Estado { get; set; }
        public string Archivo { get; set; }
        public int Porcentaje { get; set; }
        public bool Indeterminado { get; set; }
        public bool Instalando { get; set; }
        public int Procesados { get; set; }
        public int Total { get; set; }
        public int Nuevos { get; set; }
        public int Actualizados { get; set; }
        public int SinCambios { get; set; }
        public int Omitidos { get; set; }
        public bool ModoInstalacion { get; set; }
    }
}
