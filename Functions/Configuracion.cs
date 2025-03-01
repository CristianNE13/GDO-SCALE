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
        public string PuertoBascula2 { get; set; }
        public int BaudRateBascula12 { get; set; }
        public string ParityBascula12 { get; set; }
        public string StopBitsBascula12 { get; set; }
        public int DataBitsBascula12 { get; set; }
        public int EntradaSensorUnitaria { get; set; }
        public int EntradaSensorProp65 { get; set; }
        public int EntradaSensorMaster { get; set; }
        public int EntradaSensorSelladora { get; set; }
        public int SalidaDispensadoraUnitaria { get; set; }
        public int SalidaDispensadoraProp65 { get; set; }
        public int SalidaDispensadoraMaster { get; set; }
        public int SalidaSelladora { get; set; }
        public string PuertoSealevel { get; set; }
        public int BaudRateSea { get; set; }
        public string ZebraName { get; set; }

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