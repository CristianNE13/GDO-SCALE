using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace Scale_Program.Functions
{
    public class Bitacora
    {
        public static string NombreDelArchivo { get; } =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bitacora.xml");

        public int PreviousModeloID { get; set; }

        public List<RegistroBitacora> Registros { get; set; }

        public RegistroBitacora ObtenerRegistro(int modeloID)
        {
            if (Registros == null) Registros = new List<RegistroBitacora>();

            RegistroBitacora registro = null;

            for (var i = 0; i < Registros.Count; i++)
                if (Registros[i].ModeloID == modeloID)
                {
                    registro = Registros[i];
                    break;
                }

            if (registro == null)
            {
                registro = new RegistroBitacora
                {
                    ModeloID = modeloID,
                    Completadas = 0,
                    Rechazos = 0,
                    UltimaActualizacion = DateTime.Now
                };

                Registros.Add(registro);
            }

            return registro;
        }

        /// <summary>
        ///     Guarda una copia en la ruta especificada
        /// </summary>
        /// <param name="rutaDeArchivo"></param>
        public void Guardar(string directorio)
        {
            var archivo = Path.Combine(directorio, NombreDelArchivo);

            var serializer = new XmlSerializer(typeof(Bitacora));
            using (var sw = new StreamWriter(archivo))
            {
                serializer.Serialize(sw.BaseStream, this);
            }
        }

        /// <summary>
        ///     Guarda la configuracion actual en la misma ruta que se utilizo para cargar <see cref="Guardar" />
        /// </summary>
        public static Bitacora Cargar(string directorio)
        {
            Bitacora bitacora = null;

            var archivo = Path.Combine(directorio, NombreDelArchivo);

            if (File.Exists(archivo))
                try
                {
                    var serializer = new XmlSerializer(typeof(Bitacora));
                    using (var sr = new StreamReader(archivo))
                    {
                        bitacora = serializer.Deserialize(sr.BaseStream) as Bitacora;
                    }
                }
                catch
                {
                    // an exception is thrown when the file is corrupted or the xml
                    // is not well formed, just ignore the exception.
                    bitacora = new Bitacora();
                }
            else
                bitacora = new Bitacora();

            return bitacora;
        }
    }

    public class RegistroBitacora
    {
        public int ModeloID { get; set; }

        public int Completadas { get; set; }

        public int Rechazos { get; set; }

        public DateTime UltimaActualizacion { get; set; }
    }
}