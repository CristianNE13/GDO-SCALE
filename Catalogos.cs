using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;
using Scale_Program.Functions;

namespace Scale_Program
{
    public partial class Catalogos : Window
    {
        public readonly string filePathExcel = "CatalogosData.xlsx";

        public Catalogos()
        {
            InitializeComponent();

            if (!File.Exists(filePathExcel)) CrearArchivoBase();

            CargarDatosModelos();

            CargarDatosArticulos();

            NoModeloTBox.Focus();
        }

        public ObservableCollection<Modelo> Modelos { get; set; }
        public ObservableCollection<Articulo> Articulos { get; set; }

        private void CrearArchivoBase()
        {
            using (var workbook = new XLWorkbook())
            {
                var modelosSheet = workbook.Worksheets.Add("Modelos");
                modelosSheet.Cell(1, 1).Value = "NoModelo";
                modelosSheet.Cell(1, 2).Value = "ModProceso";
                modelosSheet.Cell(1, 3).Value = "Descripcion";
                modelosSheet.Cell(1, 4).Value = "UsaSensor1";
                modelosSheet.Cell(1, 5).Value = "UsaSensor2";
                modelosSheet.Cell(1, 6).Value = "Activo";

                var articulosSheet = workbook.Worksheets.Add("Articulos");
                articulosSheet.Cell(1, 1).Value = "NoParte";
                articulosSheet.Cell(1, 2).Value = "ModProceso";
                articulosSheet.Cell(1, 3).Value = "Proceso";
                articulosSheet.Cell(1, 4).Value = "Paso";
                articulosSheet.Cell(1, 5).Value = "Descripcion";
                articulosSheet.Cell(1, 6).Value = "PesoMin";
                articulosSheet.Cell(1, 7).Value = "PesoMax";
                articulosSheet.Cell(1, 8).Value = "Cantidad";
                articulosSheet.Cell(1, 9).Value = "Tag";

                var completadoSheet = workbook.Worksheets.Add("Completados");
                completadoSheet.Cell(1, 1).Value = "Fecha";
                completadoSheet.Cell(1, 2).Value = "NoParte";
                completadoSheet.Cell(1, 3).Value = "ModModelo";
                completadoSheet.Cell(1, 4).Value = "Proceso";
                completadoSheet.Cell(1, 5).Value = "PesoDetectado";
                completadoSheet.Cell(1, 6).Value = "Estado";
                completadoSheet.Cell(1, 7).Value = "Tag";

                workbook.SaveAs(filePathExcel);
            }
        }

        private void CargarDatosModelos()
        {
            using (var workbook = new XLWorkbook(filePathExcel))
            {
                var worksheet = workbook.Worksheet("Modelos");
                Modelos = new ObservableCollection<Modelo>(
                    worksheet.RowsUsed()
                        .Skip(1)
                        .Select(row => new Modelo
                        {
                            NoModelo = row.Cell(1).GetValue<string>(),
                            ModProceso = row.Cell(2).GetValue<int>(),
                            Descripcion = row.Cell(3).GetValue<string>(),
                            UsaSensor1 = row.Cell(4).GetValue<bool>(),
                            UsaSensor2 = row.Cell(5).GetValue<bool>(),
                            Activo = row.Cell(6).GetValue<bool>()
                        })
                );
                ModeloDataGrid.ItemsSource = Modelos;
            }
        }

        public void CargarDatosArticulos()
        {
            using (var workbook = new XLWorkbook(filePathExcel))
            {
                var worksheet = workbook.Worksheet("Articulos");
                Articulos = new ObservableCollection<Articulo>(
                    worksheet.RowsUsed()
                        .Skip(1)
                        .Select(row => new Articulo
                        {
                            NoParte = row.Cell(1).GetValue<string>(),
                            ModProceso = row.Cell(2).GetValue<int>(),
                            Proceso = row.Cell(3).GetValue<int>(),
                            Paso = row.Cell(4).GetValue<int>(),
                            Descripcion = row.Cell(5).GetValue<string>(),
                            PesoMin = row.Cell(6).GetValue<double>(),
                            PesoMax = row.Cell(7).GetValue<double>(),
                            Cantidad = row.Cell(8).GetValue<int>()
                        })
                );
                ArticuloDataGrid.ItemsSource = Articulos;
            }
        }

        private void ModeloAgregarBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(NoModeloTBox.Text) ||
                    string.IsNullOrWhiteSpace(ModeloDescripcionTBox.Text) ||
                    string.IsNullOrWhiteSpace(tbxProcesoModelo.Text))
                {
                    MessageBox.Show("Por favor, llena todos los campos antes de agregar.", "Campos Vacíos");
                    return;
                }

                if (!int.TryParse(tbxProcesoModelo.Text, out var proceso))
                {
                    MessageBox.Show("El proceso no es válido.");
                    return;
                }

                var modelo = new Modelo
                {
                    NoModelo = NoModeloTBox.Text,
                    ModProceso = proceso,
                    Descripcion = ModeloDescripcionTBox.Text,
                    UsaSensor1 = false,
                    UsaSensor2 = false,
                    Activo = true
                };

                Modelos.Add(modelo);

                using (var workbook = new XLWorkbook(filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Modelos");
                    var lastRow = worksheet.LastRowUsed().RowNumber();
                    var newRow = lastRow + 1;

                    worksheet.Cell(newRow, 1).Value = modelo.NoModelo;
                    worksheet.Cell(newRow, 2).Value = modelo.ModProceso;
                    worksheet.Cell(newRow, 3).Value = modelo.Descripcion;
                    worksheet.Cell(newRow, 4).Value = modelo.UsaSensor1;
                    worksheet.Cell(newRow, 5).Value = modelo.UsaSensor2;
                    worksheet.Cell(newRow, 6).Value = modelo.Activo;

                    workbook.Save();
                }

                NoModeloTBox.Clear();
                ModeloDescripcionTBox.Clear();
                tbxProcesoModelo.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al agregar modelo: {ex.Message}");
            }
        }

        private void ModeloEliminarBtn_Click(object sender, RoutedEventArgs e)
        {
            var selectedModelo = ModeloDataGrid.SelectedItem as Modelo;

            if (selectedModelo == null)
            {
                MessageBox.Show("Por favor, selecciona un modelo para eliminar.");
                return;
            }

            Modelos.Remove(selectedModelo);

            using (var workbook = new XLWorkbook(filePathExcel))
            {
                var worksheet = workbook.Worksheet("Modelos");
                worksheet.Clear();

                worksheet.Cell(1, 1).Value = "NoModelo";
                worksheet.Cell(1, 2).Value = "Proceso";
                worksheet.Cell(1, 3).Value = "Descripcion";
                worksheet.Cell(1, 4).Value = "UsaSensor1";
                worksheet.Cell(1, 5).Value = "UsaSensor2";
                worksheet.Cell(1, 6).Value = "Activo";

                var row = 2;
                foreach (var modelo in Modelos)
                {
                    worksheet.Cell(row, 1).Value = modelo.NoModelo;
                    worksheet.Cell(row, 2).Value = modelo.ModProceso;
                    worksheet.Cell(row, 3).Value = modelo.Descripcion;
                    worksheet.Cell(row, 4).Value = modelo.UsaSensor1;
                    worksheet.Cell(row, 5).Value = modelo.UsaSensor2;
                    worksheet.Cell(row, 6).Value = modelo.Activo;
                    row++;
                }

                workbook.Save();
            }
        }

        private void GuardarArticulosEnExcel()
        {
            try
            {
                using (var workbook = new XLWorkbook(filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Articulos");
                    worksheet.Clear();

                    worksheet.Cell(1, 1).Value = "NoParte";
                    worksheet.Cell(1, 2).Value = "ModProceso";
                    worksheet.Cell(1, 3).Value = "Proceso";
                    worksheet.Cell(1, 4).Value = "Paso";
                    worksheet.Cell(1, 5).Value = "Descripcion";
                    worksheet.Cell(1, 6).Value = "PesoMin";
                    worksheet.Cell(1, 7).Value = "PesoMax";
                    worksheet.Cell(1, 8).Value = "Cantidad";

                    var row = 2;
                    foreach (var articulo in Articulos)
                    {
                        worksheet.Cell(row, 1).Value = articulo.NoParte;
                        worksheet.Cell(row, 2).Value = articulo.ModProceso;
                        worksheet.Cell(row, 3).Value = articulo.Proceso;
                        worksheet.Cell(row, 4).Value = articulo.Paso;
                        worksheet.Cell(row, 5).Value = articulo.Descripcion;
                        worksheet.Cell(row, 6).Value = articulo.PesoMin;
                        worksheet.Cell(row, 7).Value = articulo.PesoMax;
                        worksheet.Cell(row, 8).Value = articulo.Cantidad;
                        row++;
                    }

                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar los artículos en el archivo Excel: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ArticuloNuevoBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(NoArticuloTbox.Text) ||
                    string.IsNullOrWhiteSpace(ArticuloDescripcionTbox.Text) ||
                    string.IsNullOrWhiteSpace(ArticuloPesoMinTbox.Text) ||
                    string.IsNullOrWhiteSpace(ArticuloPesoMaxTbox.Text) ||
                    string.IsNullOrWhiteSpace(txbProceso.Text))
                {
                    MessageBox.Show("Por favor, llena todos los campos antes de agregar.", "Campos Vacíos",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!int.TryParse(txbProceso.Text, out var proceso) ||
                    !int.TryParse(txbPaso.Text, out var paso))
                {
                    MessageBox.Show("El Proceso o Paso no son válidos.", "Error", MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                if (!int.TryParse(txbModProceso.Text, out var modproceso) ||
                    !int.TryParse(txbCantidad.Text, out var cantidad))
                {
                    MessageBox.Show("El ModProceso", "Error", MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                if (!double.TryParse(ArticuloPesoMinTbox.Text, out var pesoMin) ||
                    !double.TryParse(ArticuloPesoMaxTbox.Text, out var pesoMax))
                {
                    MessageBox.Show("Los valores de Peso Mínimo o Máximo no son válidos.", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var articulo = new Articulo
                {
                    NoParte = NoArticuloTbox.Text,
                    ModProceso = modproceso,
                    Proceso = proceso,
                    Paso = paso,
                    Descripcion = ArticuloDescripcionTbox.Text,
                    PesoMin = pesoMin,
                    PesoMax = pesoMax,
                    Cantidad = cantidad
                };

                using (var workbook = new XLWorkbook(filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Articulos");
                    var lastRow = worksheet.LastRowUsed().RowNumber();
                    var newRow = lastRow + 1;

                    worksheet.Cell(newRow, 1).Value = articulo.NoParte;
                    worksheet.Cell(newRow, 2).Value = articulo.ModProceso;
                    worksheet.Cell(newRow, 3).Value = articulo.Proceso;
                    worksheet.Cell(newRow, 4).Value = articulo.Paso;
                    worksheet.Cell(newRow, 5).Value = articulo.Descripcion;
                    worksheet.Cell(newRow, 6).Value = articulo.PesoMin;
                    worksheet.Cell(newRow, 7).Value = articulo.PesoMax;
                    worksheet.Cell(newRow, 8).Value = articulo.Cantidad;

                    workbook.Save();
                }

                MessageBox.Show("Artículo agregado correctamente.", "Éxito", MessageBoxButton.OK,
                    MessageBoxImage.Information);

                NoArticuloTbox.Clear();
                ArticuloDescripcionTbox.Clear();
                ArticuloPesoMinTbox.Clear();
                ArticuloPesoMaxTbox.Clear();
                txbProceso.Clear();
                txbPaso.Clear();

                CargarDatosArticulos();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al agregar artículo: {ex.Message}", "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void ArticuloEliminarBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedArticulo = ArticuloDataGrid.SelectedItem as Articulo;

                if (selectedArticulo == null)
                {
                    MessageBox.Show("Por favor, selecciona un artículo para eliminar.", "Selección Vacía",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!File.Exists(filePathExcel))
                {
                    MessageBox.Show("El archivo Excel no existe.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                using (var workbook = new XLWorkbook(filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Articulos");
                    var rows = worksheet.RowsUsed().Skip(1);

                    var rowToDelete = rows.FirstOrDefault(row =>
                        row.Cell(1).GetValue<string>() == selectedArticulo.NoParte &&
                        row.Cell(2).GetValue<int>() == selectedArticulo.ModProceso &&
                        row.Cell(3).GetValue<int>() == selectedArticulo.Proceso &&
                        row.Cell(4).GetValue<int>() == selectedArticulo.Paso &&
                        row.Cell(5).GetValue<string>() == selectedArticulo.Descripcion &&
                        Math.Abs(row.Cell(6).GetValue<double>() - selectedArticulo.PesoMin) < 0.0001 &&
                        Math.Abs(row.Cell(7).GetValue<double>() - selectedArticulo.PesoMax) < 0.0001 &&
                        row.Cell(8).GetValue<int>() == selectedArticulo.Cantidad);

                    if (rowToDelete != null)
                    {
                        rowToDelete.Delete();
                        workbook.Save();

                        MessageBox.Show("Artículo eliminado correctamente.", "Éxito", MessageBoxButton.OK,
                            MessageBoxImage.Information);

                        CargarDatosArticulos();
                    }
                    else
                    {
                        MessageBox.Show("No se encontró el artículo en el archivo Excel.", "Error", MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al eliminar el artículo: {ex.Message}", "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void GuardarModelosEnExcel()
        {
            try
            {
                using (var workbook = new XLWorkbook(filePathExcel))
                {
                    var worksheet = workbook.Worksheet("Modelos");
                    worksheet.Clear();

                    worksheet.Cell(1, 1).Value = "NoModelo";
                    worksheet.Cell(1, 2).Value = "Proceso";
                    worksheet.Cell(1, 3).Value = "Descripcion";
                    worksheet.Cell(1, 4).Value = "UsaSensor1";
                    worksheet.Cell(1, 5).Value = "UsaSensor2";
                    worksheet.Cell(1, 6).Value = "Activo";

                    var row = 2;
                    foreach (var modelo in Modelos)
                    {
                        worksheet.Cell(row, 1).Value = modelo.NoModelo;
                        worksheet.Cell(row, 2).Value = modelo.ModProceso;
                        worksheet.Cell(row, 3).Value = modelo.Descripcion;
                        worksheet.Cell(row, 4).Value = modelo.UsaSensor1;
                        worksheet.Cell(row, 5).Value = modelo.UsaSensor2;
                        worksheet.Cell(row, 6).Value = modelo.Activo;
                        row++;
                    }

                    workbook.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar los modelos en el archivo Excel: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnGuardarCambios_Click_1(object sender, RoutedEventArgs e)
        {
            GuardarArticulosEnExcel();

            MessageBox.Show("Cambios guardados correctamente en el archivo Excel.", "Éxito", MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void BtnGuardarModelos_Click(object sender, RoutedEventArgs e)
        {
            GuardarModelosEnExcel();

            MessageBox.Show("Cambios guardados correctamente en el archivo Excel.", "Éxito", MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}