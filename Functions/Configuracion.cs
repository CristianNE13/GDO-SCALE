using System;
using System.IO;
using System.Xml.Serialization;

namespace Scale_Program.Functions
{
    [Serializable]
    [XmlRoot("Configuracion")]
    public class Configuracion
    {
        public string PuertoBascula1 { get; set; }
        public int BaudRateBascula12 { get; set; }
        public string ParityBascula12 { get; set; }
        public string StopBitsBascula12 { get; set; }
        public int DataBitsBascula12 { get; set; }
        public string PuertoSealevel { get; set; }
        public int BaudRateSea { get; set; }
        public string User { get; set; }
        public int InputPick2L0 { get; set; }
        public int OutputPick2L0 { get; set; }
        public int InputBoton { get; set; }
        public bool CheckShutOff { get; set; }
        public int ShutOff { get; set; }
        public int Piston { get; set; }
        public string IpCamara { get; set; }
        public int PuertoCamara { get; set; }
        public string BasculaMarca { get; set; }
        public bool CheckSealevelEthernet { get; set; }
        public string SealevelIP { get; set; }
        public int InputSensor0 { get; set; }
        public int InputSensor1 { get; set; }

        public static string RutaArchivoConf { get; } =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "configuration.xml");

        /// <summary>
        ///     Carga el archivo de configuración desde la ruta especificada.
        /// </summary>
        public static Configuracion Cargar(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"No se encontró el archivo de configuración: {filePath}");

            var serializer = new XmlSerializer(typeof(Configuracion));
            using (var sr = new StreamReader(filePath))
            {
                return (Configuracion)serializer.Deserialize(sr);
            }
        }

        /// <summary>
        ///     Guarda la configuración actual en la ruta especificada.
        /// </summary>
        public void Guardar(string rutaDeArchivo)
        {
            var serializer = new XmlSerializer(typeof(Configuracion));
            using (var sw = new StreamWriter(rutaDeArchivo))
            {
                serializer.Serialize(sw, this);
            }
        }
    }
}