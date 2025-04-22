using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Scale_Program.Functions;

namespace Scale_Program
{
    public partial class Catalogos : Window
    {
        public event Action CambiosGuardados;
        public ObservableCollection<Modelo> Modelos { get; set; }
        public ObservableCollection<Articulo> Articulos { get; set; }

        public Catalogos()
        {
            InitializeComponent();

            CargarDatosModelos();

            CargarDatosArticulos();

            NoModeloTBox.Focus();
        }

        private void CargarDatosModelos()
        {
            using (var db = new dc_missingpartsEntities())
            {
                var modelosDb = db.Modelos
                    .OrderBy(m => m.ModProceso)
                    .ToList();
                
                Modelos = new ObservableCollection<Modelo>(modelosDb);
                
                ModeloDataGrid.ItemsSource = Modelos;
            }
        }

        public void CargarDatosArticulos()
        {

            using (var db = new dc_missingpartsEntities())
            {
                var modelosDb = db.Modelos
                    .Where(m => m.Activo)
                    .Select(m => m.ModProceso)
                    .ToList();

                var articulosDb = db.Articulos
                    .Where(articulo => modelosDb.Contains(articulo.ModProceso))
                    .OrderBy(m => m.ModProceso)
                    .ThenBy(p => p.Proceso)
                    .ThenBy(s => s.Paso)
                    .ToList();

                Articulos = new ObservableCollection<Articulo>(articulosDb);

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

                if (ckb_Bascula2.IsChecked != null && !ckb_Bascula1.IsChecked.Value && !ckb_Bascula2.IsChecked.Value)
                {
                    MessageBox.Show("Debes utilizar una bascula.", "Seleccionar bascula");
                    return;
                }

                if (ckb_ConteoCajas.IsChecked != null && ckb_Bascula2.IsChecked != null &&
                    ckb_Bascula1.IsChecked != null && ckb_Bascula1.IsChecked.Value && !ckb_Bascula2.IsChecked.Value &&
                    ckb_ConteoCajas.IsChecked.Value)
                {
                    MessageBox.Show("No puedes utilizar una conteo cajas, necesitas utilizar bascula 1 y 2.",
                        "Conteo cajas");
                    return;
                }

                var modelo = new Modelo
                {
                    NoModelo = NoModeloTBox.Text,
                    ModProceso = proceso,
                    Descripcion = ModeloDescripcionTBox.Text,
                    UsaBascula1 = ckb_Bascula1.IsChecked != null && ckb_Bascula1.IsChecked.Value,
                    UsaBascula2 = ckb_Bascula2.IsChecked != null && ckb_Bascula2.IsChecked.Value,
                    UsaConteoCajas = ckb_ConteoCajas.IsChecked != null && ckb_ConteoCajas.IsChecked.Value,
                    CantidadCajas = int.TryParse(txb_CantidadCajas.Text, out var cantidad) ? cantidad : 0,
                    Etapa1 = txb_Etapa1.Text,
                    Etapa2 = ckb_Bascula1.IsChecked != null && ckb_Bascula1.IsChecked.Value ? txb_Etapa2.Text : txb_Etapa1Bascula2.Text,
                    Activo = true
                };

                Modelos.Add(modelo);

                using (var db = new dc_missingpartsEntities())
                {
                    db.Modelos.Add(modelo);
                    db.SaveChanges();
                }

                NoModeloTBox.Clear();
                ModeloDescripcionTBox.Clear();
                tbxProcesoModelo.Clear();
                ckb_Bascula1.IsChecked = false;
                ckb_Bascula2.IsChecked = false;
                ckb_ConteoCajas.IsChecked = false;
                txb_CantidadCajas.Clear();
                txb_Etapa1Bascula2.Clear();
                txb_Etapa1.Clear();
                txb_Etapa2.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al agregar modelo: {ex.Message}");
            }
        }

        private void ModeloEliminarBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(ModeloDataGrid.SelectedItem is Modelo selectedModelo))
                {
                    MessageBox.Show("Por favor, selecciona un modelo para eliminar.");
                    return;
                }

                Modelos.Remove(selectedModelo);

                using (var db = new dc_missingpartsEntities())
                {
                    var modeloDb = db.Modelos.Find(selectedModelo.Id);
                    if (modeloDb == null)
                    {
                        MessageBox.Show("Modelo no encontrado en la base de datos.");
                        return;
                    }

                    db.Modelos.Remove(modeloDb);
                    db.SaveChanges();
                }
                CargarDatosModelos();
                MessageBox.Show("Modelo eliminado correctamente.", "Éxito", MessageBoxButton.OK,
                                       MessageBoxImage.Information);
            }
            catch (Exception exception)
            {
                MessageBox.Show($"Error al eliminar el modelo: {exception.Message}");
                throw;
            }

        }

        private void GuardarArticulos()
        {
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var idsEnBD = db.Articulos.Select(a => a.Id).ToList();
                    var idsEnPantalla = Articulos.Select(a => a.Id).ToList();

                    var idsAEliminar = idsEnBD.Except(idsEnPantalla).ToList();

                    foreach (var id in idsAEliminar)
                    {
                        var articuloEliminar = db.Articulos.FirstOrDefault(a => a.Id == id);
                        if (articuloEliminar != null)
                            db.Articulos.Remove(articuloEliminar);
                    }

                    foreach (var articuloLocal in Articulos)
                    {
                        var articuloBD = db.Articulos.FirstOrDefault(a => a.Id == articuloLocal.Id);

                        if (articuloBD != null)
                        {
                            articuloBD.NoParte = articuloLocal.NoParte;
                            articuloBD.ModProceso = articuloLocal.ModProceso;
                            articuloBD.Proceso = articuloLocal.Proceso;
                            articuloBD.Paso = articuloLocal.Paso;
                            articuloBD.Descripcion = articuloLocal.Descripcion;
                            articuloBD.PesoMin = articuloLocal.PesoMin;
                            articuloBD.PesoMax = articuloLocal.PesoMax;
                            articuloBD.Cantidad = articuloLocal.Cantidad;
                        }
                        else
                            db.Articulos.Add(articuloLocal);
                    }

                    db.SaveChanges();
                }

                MessageBox.Show("Los artículos se guardaron correctamente en la base de datos.", "Éxito",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar los artículos: {ex.Message}", "Error",
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
                    string.IsNullOrWhiteSpace(txbCantidad.Text) ||
                    string.IsNullOrWhiteSpace(txbPaso.Text) ||
                    string.IsNullOrWhiteSpace(txbModProceso.Text))
                {
                    MessageBox.Show("Por favor, llena todos los campos antes de agregar.", "Campos Vacíos",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!int.TryParse(txbProceso.Text, out var proceso) ||
                    !int.TryParse(txbPaso.Text, out var paso) || paso == 0)
                {
                    MessageBox.Show("El Proceso o paso no son válidos.", "Error", MessageBoxButton.OK,
                        MessageBoxImage.Error);
                    return;
                }

                if (!int.TryParse(txbModProceso.Text, out var modproceso) ||
                    !int.TryParse(txbCantidad.Text, out var cantidad))
                {
                    MessageBox.Show("El ModProceso o cantidad no son validos", "Error", MessageBoxButton.OK,
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

                if (pesoMin > pesoMax)
                {
                    MessageBox.Show("El Peso Mínimo no puede ser mayor que el Peso Máximo.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
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

                using (var db = new dc_missingpartsEntities())
                {
                    db.Articulos.Add(articulo);
                    db.SaveChanges();
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

                using (var db = new dc_missingpartsEntities())
                {
                    var articuloDb = db.Articulos.Find(selectedArticulo.Id);
                    if (articuloDb != null)
                    {
                        db.Articulos.Remove(articuloDb);
                        db.SaveChanges();
                    }
                
                }
                CargarDatosArticulos();
                MessageBox.Show("Artículo eliminado correctamente.", "Éxito", MessageBoxButton.OK,
                                       MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al eliminar el artículo: {ex.Message}", "Error", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void GuardarModelos()
        {
            try
            {
                using (var db = new dc_missingpartsEntities())
                {
                    var modelosBD = db.Modelos.ToList();
                    var idsPantalla = Modelos.Select(m => m.Id).ToHashSet();

                    var modelosAEliminar = modelosBD.Where(m => !idsPantalla.Contains(m.Id)).ToList();
                    db.Modelos.RemoveRange(modelosAEliminar);

                    foreach (var modeloLocal in Modelos)
                    {
                        var modeloBD = modelosBD.FirstOrDefault(m => m.Id == modeloLocal.Id);
                        if (modeloBD != null)
                        {
                            modeloBD.NoModelo = modeloLocal.NoModelo;
                            modeloBD.Descripcion = modeloLocal.Descripcion;
                            modeloBD.UsaBascula1 = modeloLocal.UsaBascula1;
                            modeloBD.UsaBascula2 = modeloLocal.UsaBascula2;
                            modeloBD.UsaConteoCajas = modeloLocal.UsaConteoCajas;
                            modeloBD.CantidadCajas = modeloLocal.CantidadCajas;
                            modeloBD.Etapa1 = modeloLocal.Etapa1;
                            modeloBD.Etapa2 = modeloLocal.Etapa2;
                            modeloBD.Activo = modeloLocal.Activo;
                        }
                        else
                            db.Modelos.Add(modeloLocal);
                    }

                    db.SaveChanges();
                }


                MessageBox.Show("Los cambios se guardaron correctamente en la base de datos.", "Éxito",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar los modelos en la base de datos: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void BtnGuardarModelos_Click(object sender, RoutedEventArgs e)
        {
            GuardarModelos();
            CambiosGuardados?.Invoke();
            CargarDatosArticulos();
        }

        private void ckb_Bascula1_Checked(object sender, RoutedEventArgs e)
        {
            if (ckb_Bascula1.IsChecked.Value && ckb_Bascula2.IsChecked.Value)
            {
                lblConteoCajas.Visibility = Visibility.Visible;
                ckb_ConteoCajas.Visibility = Visibility.Visible;

                lblUltimoEtapa2.Visibility = Visibility.Hidden;
                txb_Etapa1Bascula2.Visibility = Visibility.Hidden;

                lblPrimerEtapa2.Visibility = Visibility.Visible;
                txb_Etapa2.Visibility = Visibility.Visible;
            }

            lblUltimoEtapa1.Visibility = Visibility.Visible;
            txb_Etapa1.Visibility = Visibility.Visible;
        }

        private void ckb_Bascula1_Unchecked(object sender, RoutedEventArgs e)
        {
            if (!ckb_Bascula1.IsChecked.Value && ckb_Bascula2.IsChecked.Value)
            {
                lblUltimoEtapa2.Visibility = Visibility.Visible;
                txb_Etapa1Bascula2.Visibility = Visibility.Visible;

                lblPrimerEtapa2.Visibility = Visibility.Hidden;
                txb_Etapa2.Visibility = Visibility.Hidden;
            }

            lblConteoCajas.Visibility = Visibility.Hidden;
            ckb_ConteoCajas.Visibility = Visibility.Hidden;

            lblConteoCajas.Visibility = Visibility.Hidden;
            lblCantidad.Visibility = Visibility.Hidden;
            txb_CantidadCajas.Text = "0";
            txb_CantidadCajas.Visibility = Visibility.Hidden;
            ckb_ConteoCajas.Visibility = Visibility.Hidden;
            ckb_ConteoCajas.IsChecked = false;

            lblUltimoEtapa1.Visibility = Visibility.Hidden;
            txb_Etapa1.Visibility = Visibility.Hidden;
        }

        private void ckb_ConteoCajas_Checked(object sender, RoutedEventArgs e)
        {
            lblCantidad.Visibility = Visibility.Visible;
            txb_CantidadCajas.Visibility = Visibility.Visible;
        }

        private void ckb_ConteoCajas_Unchecked(object sender, RoutedEventArgs e)
        {
            lblCantidad.Visibility = Visibility.Hidden;
            txb_CantidadCajas.Visibility = Visibility.Hidden;
        }

        private void ckb_Bascula2_Checked(object sender, RoutedEventArgs e)
        {
            if (ckb_Bascula1.IsChecked.Value)
            {
                lblPrimerEtapa2.Visibility = Visibility.Visible;
                txb_Etapa2.Visibility = Visibility.Visible;
            }

            if (ckb_Bascula1.IsChecked.Value && ckb_Bascula2.IsChecked.Value)
            {
                lblConteoCajas.Visibility = Visibility.Visible;
                ckb_ConteoCajas.Visibility = Visibility.Visible;
            }
            else
            {
                lblUltimoEtapa2.Visibility = Visibility.Visible;
                txb_Etapa1Bascula2.Visibility = Visibility.Visible;
            }
        }

        private void ckb_Bascula2_Unchecked(object sender, RoutedEventArgs e)
        {
            lblUltimoEtapa2.Visibility = Visibility.Hidden;
            txb_Etapa2.Visibility = Visibility.Hidden;

            lblConteoCajas.Visibility = Visibility.Hidden;
            lblCantidad.Visibility = Visibility.Hidden;

            txb_CantidadCajas.Text = "0";
            txb_CantidadCajas.Visibility = Visibility.Hidden;
            ckb_ConteoCajas.Visibility = Visibility.Hidden;
            ckb_ConteoCajas.IsChecked = false;

            lblPrimerEtapa2.Visibility = Visibility.Hidden;
            txb_Etapa1Bascula2.Visibility = Visibility.Hidden;
        }

        private void btnGuardarArticulos_Click(object sender, RoutedEventArgs e)
        {
            GuardarArticulos();
            CargarDatosModelos();
        }
    }
}