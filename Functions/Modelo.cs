namespace Scale_Program.Functions
{
    public class Modelo
    {
        public string NoModelo { get; set; }
        public int ModProceso { get; set; }
        public string Descripcion { get; set; }
        public bool UsaBascula1 { get; set; }
        public bool UsaBascula2 { get; set; }
        public bool UsaConteoCajas { get; set; }
        public int CantidadCajas { get; set; }
        public string Etapa1 { get; set; }
        public string Etapa2 { get; set; }
        public bool Activo { get; set; }

        public Modelo()
        {
            NoModelo = string.Empty;
            ModProceso = 0;
            Descripcion = string.Empty;
            UsaBascula1 = false;
            UsaBascula2 = false;
            UsaConteoCajas = false;
            CantidadCajas = 0;
            Etapa1 = string.Empty;
            Etapa2 = string.Empty;
            Activo = true;
        }

        public override string ToString()
        {
            return NoModelo;
        }
    }
}