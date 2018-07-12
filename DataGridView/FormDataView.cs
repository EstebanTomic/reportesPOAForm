using iTextSharp.text;
using iTextSharp.text.pdf;
using MySql.Data.MySqlClient;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace DataGridView
{
    public partial class FormDataView : Form
    {
        // Variable global que almacenará el canal seleccionado en formulario padre
        private String canal;

        public FormDataView()
        {
            InitializeComponent();
        }

        private void FormDataView_Load(object sender, EventArgs e)
        {
            // Se debe obtener CANAL desde formulario padre
            //this.canal = "CON";
            this.canal = "IND";
            // Validar si es necesaria una llamada a BD

            // Llamamos a método que se encarga de armar lógica de Combos
            ArmaCombosInformes(this.canal);

        }

        #region Consultas Base de Datos

        private void DataInformesBD(String numeroInforme)
        {
            // Se crea la conexión a la base de datos
            MySqlConnection _conexion = new MySqlConnection("server=localhost;user id=root;password=root;persistsecurityinfo=True;database=test;SslMode=none");

            // Se abre la conexion
            _conexion.Open();

            // Se crea un DataTable que almacenará los datos desde donde se cargaran los datos al DataGridView
            DataTable dtDatos = new DataTable();

            // Se crea un MySqlAdapter para obtener los datos de la base
            MySqlDataAdapter mdaDatos;
            if (numeroInforme.Equals("1"))
            {
                mdaDatos = new MySqlDataAdapter("SELECT producto, ventaold, ventaant, poa, acumenesep, ventaantold, poaactant FROM informe1 ORDER BY ID ASC", _conexion);
                // Con la información del adaptador se rellena el DataTable
                mdaDatos.Fill(dtDatos);
            }
            else if (numeroInforme.Equals("2"))
            {
                mdaDatos = new MySqlDataAdapter("select * From informe2", _conexion);
                // Con la información del adaptador se rellena el DataTable
                mdaDatos.Fill(dtDatos);
            }
            // ARMAR QUERY
            else if(numeroInforme.Equals("3"))
            {
                mdaDatos = new MySqlDataAdapter("SELECT producto, planteorico, efectos, planaccion, plandef, efcom, pacc FROM informe3 A order by Id ASC", _conexion);
                // Con la información del adaptador se rellena el DataTable
                mdaDatos.Fill(dtDatos);
            }
            else if (numeroInforme.Equals("4"))
            {
                mdaDatos = new MySqlDataAdapter("SELECT producto, ene, feb, mar, abr, may, jun, jul, ago, sep, oct, nov, dic, total FROM informe4 ORDER BY id ASC", _conexion);
                // Con la información del adaptador se rellena el DataTable
                mdaDatos.Fill(dtDatos);
            }
            else if (numeroInforme.Equals("5"))
            {
                mdaDatos = new MySqlDataAdapter("SELECT producto, precio, prodser, fventa, prompub, operac, total FROM informe5 ORDER BY ID ASC", _conexion);
                // Con la información del adaptador se rellena el DataTable
                mdaDatos.Fill(dtDatos);
            }
            else if (numeroInforme.Equals("7"))
            {
                mdaDatos = new MySqlDataAdapter("SELECT * FROM informe7", _conexion);
                // Con la información del adaptador se rellena el DataTable
                mdaDatos.Fill(dtDatos);
            }

            // Se asigna el DataTable como origen de datos del DataGridView
            dataGridView1.DataSource = dtDatos;

            // Se cierra la conexión a la base de datos
            _conexion.Close();
        }

        private void DatosBD(String numeroInforme)
        {
            // Se crea la conexión a la base de datos
            MySqlConnection _conexion = new MySqlConnection("server=localhost;user id=root;password=root;persistsecurityinfo=True;database=test;SslMode=none");

            // Se abre la conexion
            _conexion.Open();

            // Se crea un DataTable que almacenará los datos desde donde se cargaran los datos al DataGridView
            DataTable dtDatos = new DataTable();

            // Se crea un MySqlAdapter para obtener los datos de la base
            MySqlDataAdapter mdaDatos = new MySqlDataAdapter("SELECT * FROM prueba ORDER BY ID DESC", _conexion);

            // Con la información del adaptador se rellena el DataTable
            mdaDatos.Fill(dtDatos);

            // Se asigna el DataTable como origen de datos del DataGridView
            dataGridView1.DataSource = dtDatos;

            // Se cierra la conexión a la base de datos
            _conexion.Close();
        }
        #endregion

        #region Armado de ComboBox
        private void ArmaCombosInformes(String canal) {

            // Seteo de combobox
            // Armamos Combobox de Variables
            this.comboBoxEntidad.SelectedIndex = 0;
            this.comboBoxEntidad.DropDownStyle = ComboBoxStyle.DropDownList;
            this.comboBoxVariable.Items.Clear();

            // Permisos según los descritos en BD (PA_PARAMETROS en LOCAL)
            String permisos = "";
            if (canal.Equals("CON") || canal.Equals("CUP") || canal.Equals("SOL") || canal.Equals("TCT"))
                permisos = "V";
            if (canal.Equals("IND") || canal.Equals("LUB") || canal.Equals("LUE"))
                permisos = "VMC";

            // Armamos Combobox de Variables
            if (permisos.IndexOf("V") != -1)
                this.comboBoxVariable.Items.AddRange(new object[] { "1. Volumen" });
            if (permisos.IndexOf("M") != -1)
                this.comboBoxVariable.Items.AddRange(new object[] { "2. Márgenes" });
            if (permisos.IndexOf("C") != -1)
                this.comboBoxVariable.Items.AddRange(new object[] { "3. Contribución" });
            if (permisos.IndexOf("$") != -1)
                this.comboBoxVariable.Items.AddRange(new object[] { "4. Montos $" });
            if (permisos.IndexOf("M") != -1 || permisos.IndexOf("C") != -1)
                this.comboBoxVariable.Items.AddRange(new object[] { "5. Mixtos" });

            this.comboBoxVariable.SelectedIndex = 0;
            this.comboBoxVariable.DropDownStyle = ComboBoxStyle.DropDownList;

            //ArmaCombosTipoInformes();

            this.comboBoxProducto.SelectedIndex = 0;
            this.comboBoxProducto.DropDownStyle = ComboBoxStyle.DropDownList;
            this.comboBoxOrden.SelectedIndex = 0;
            this.comboBoxOrden.DropDownStyle = ComboBoxStyle.DropDownList;

        }

        private void ArmaComboTipoInformes()
        {
            // Armamos Combobox de Tipo de Informes
            // Obtenemos ID de Variable seleccionada
            String variable = this.comboBoxVariable.SelectedItem.ToString().Substring(0,1);

            // Reseteamos el Combobox antes de setearlo según corresponda
            this.comboBoxTipoInforme.Items.Clear();

            if (variable.Equals("1"))
            {
                this.comboBoxTipoInforme.Items.AddRange(new object[] {
                    "1. Presentación de Proyecciones de Venta",
                    "3. Presentación Plan",
                    "4. Presentación Plan Definitivo por Mes",
                    "5. Presentación Planes de Acción",
                    "6. Presentación Correc. Efect. Comerciales",
                    "7. Ranking Crecim. Cliente P.Act v/s P.Defi"
                });
            }
            else if (variable.Equals("2"))
            {
                this.comboBoxTipoInforme.Items.AddRange(new object[] {
                    "2. Presentación de Proyecciones Contribución",
                    "3. Presentación Plan",
                    "4. Presentación Plan Definitivo por Mes",
                    "7. Ranking Crecim. Cliente P.Act v/s P.Defi"
                });
            }
            else if (variable.Equals("3"))
            {
                this.comboBoxTipoInforme.Items.AddRange(new object[] {
                    "2. Presentación de Proyecciones Contribución",
                    "3. Presentación Plan",
                    "4. Presentación Plan Definitivo por Mes",
                    "7. Ranking Crecim. Cliente P.Act v/s P.Defi"
                });
            }
            else if (variable.Equals("4"))
            {
                this.comboBoxTipoInforme.Items.AddRange(new object[] {
                    "0. Próximo Poa"});
            }
            else if (variable.Equals("5"))
            {
                this.comboBoxTipoInforme.Items.AddRange(new object[] {
                    "8. Vol,Mar,Con Cliente P.Act v/s P.Def Crecimientos",
                    "9. Vol,Mar,Con Cliente P.Act v/s P.Def Diferencias"
                });
            }
            this.comboBoxTipoInforme.SelectedIndex = 0;
            this.comboBoxTipoInforme.DropDownStyle = ComboBoxStyle.DropDownList;

        }

        private void ArmaComboAgrupacion()
        {
            // HAY QUE OBTENERLO **************************************
            String canal = "CON";
            // ********************************************************
            // Armamos Combobox de Agrupación
            // Obtenemos ID de Variable seleccionada
            String variable = this.comboBoxVariable.SelectedItem.ToString().Substring(0, 1);
            // Obtenemos ID de Tipo de informe seleccionado
            String tipoInforme = this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1);

            // Reseteamos el Combobox antes de setearlo según corresponda
            this.comboBoxAgrupacion.Items.Clear();

            if (!tipoInforme.Equals("7"))
            {
                if (!tipoInforme.Equals("1"))
                {
                    this.comboBoxAgrupacion.Items.AddRange(new object[] {
                        "1. Por Producto",
                        "2. Por Rubro"
                    });
                    if (canal.Equals("IND") && variable.Equals("1"))
                    {
                        this.comboBoxAgrupacion.Items.AddRange(new object[] {
                            "4. Por Producto Canal"
                        });
                    }
                }
                else
                {
                    this.comboBoxAgrupacion.Items.AddRange(new object[] {
                        "1. Por Producto",
                        "2. Por Rubro",
                        "3. Por Vendedor"
                    });
                    if (canal.Equals("IND") && variable.Equals("1"))
                    {
                        this.comboBoxAgrupacion.Items.AddRange(new object[] {
                            "4. Por Producto Canal"
                        });
                    }
                }
            }
            else
            {
                this.comboBoxAgrupacion.Items.AddRange(new object[] {
                    "0. Por Cliente"
                });
            }
            this.comboBoxAgrupacion.SelectedIndex = 0;
            this.comboBoxAgrupacion.DropDownStyle = ComboBoxStyle.DropDownList;

        }

        private void ArmaComboOrden()
        {
            // Obtenemos ID de Tipo de informe seleccionado
            String tipoInforme = this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1);

            // Reseteamos el Combobox antes de setearlo según corresponda
            this.comboBoxOrden.Items.Clear();

            if (tipoInforme.Equals("8") || tipoInforme.Equals("9"))
            {
                this.comboBoxOrden.Items.AddRange(new object[] {
                    "1. Volumen",
                    "2. Margen",
                    "3. Contribución"
                });
                this.comboBoxOrden.Enabled = true;
            }
            else
            {
                this.comboBoxOrden.Items.AddRange(new object[] {
                    "1. Volumen"
                });
                this.comboBoxOrden.Enabled = false;
            }
            this.comboBoxOrden.SelectedIndex = 0;
        }

        private void ArmaComboProducto()
        {
            // Obtenemos ID de Tipo de informe seleccionado
            String tipoInforme = this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1);

            // Limpiamos el Combobox
            this.comboBoxProducto.Items.Clear();
            // Validamos para que informe se debe habilitar
            if (tipoInforme.Equals("7") || tipoInforme.Equals("8") || tipoInforme.Equals("9"))
            {
                this.comboBoxProducto.Items.AddRange(new object[] {
                    "TOT. <TOTAL PRODUCTOS>",
                    "GAS. Total Gasolinas",
                    "DSL. Diesel",
                    "KER. Kerosene",
                    "GAV. Gasolina Aviacion",
                    "TUR. Turbo",
                    "PCO. Total Pet. Comb.",
                    "XTO. Taxi Amigo Total"
                });
                this.comboBoxProducto.Enabled = true;
            }
            else
            {
                this.comboBoxProducto.Items.AddRange(new object[] {
                    "TOT. <TOTAL PRODUCTOS>"
                });
                this.comboBoxProducto.Enabled = false;
            }
            this.comboBoxProducto.SelectedIndex = 0;
        }

        // Cambios en los combos que correspondan al cambiar la variable
        private void comboBoxVariable_SelectedIndexChanged(object sender, EventArgs e)
        {
            ArmaComboTipoInformes();
        }

        // Cambios en los combos que correspondan al cambiar el tipo de Informe
        private void comboBoxTipoInforme_SelectedIndexChanged(object sender, EventArgs e)
        {
            ArmaComboAgrupacion();
            ArmaComboProducto();
            ArmaComboOrden();
        }
        #endregion

        #region Generación de PDF
        private PdfPTable CreaHeader(String canal)
        {
            // Creando tabla para encabezado
            PdfPTable pdfTable1 = new PdfPTable(3);//Cantidad de columnas
            pdfTable1.WidthPercentage = 100;
            pdfTable1.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable1.DefaultCell.VerticalAlignment = Element.ALIGN_CENTER;
            pdfTable1.DefaultCell.BorderWidth = 0;

            // Se genera el nombre del canal según lo recibido como entrada
            // REVISAR
            String nombreCanal = canal.Equals("CON") ? "Concesionarios" : "Industrial";

            // Generamos datos de la tabla de encabezado
            Chunk c1 = new Chunk("Cia.Petroleos de Chile Copec S.A", FontFactory.GetFont("Microsoft Sans Serif", 10, iTextSharp.text.Font.BOLD));
            c1.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);
            Chunk c2 = new Chunk("CANAL " + nombreCanal.ToUpper(), FontFactory.GetFont("Microsoft Sans Serif", 10, iTextSharp.text.Font.BOLD));
            c2.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);
            Chunk c3 = new Chunk("Gerencia de Marketing", FontFactory.GetFont("Microsoft Sans Serif", 10, iTextSharp.text.Font.BOLD));
            c3.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);
            Chunk c4 = new Chunk("Subg. Planificación Comercial", FontFactory.GetFont("Microsoft Sans Serif", 10, iTextSharp.text.Font.BOLD));
            c4.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);
            Phrase p1 = new Phrase();
            p1.Add(c1);
            Phrase p2 = new Phrase();
            p2.Add(c2);
            Phrase p3 = new Phrase();
            p3.Add(c3);
            Phrase p4 = new Phrase();
            p4.Add(c4);
            pdfTable1.AddCell(p1);
            pdfTable1.AddCell("");
            pdfTable1.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT;
            pdfTable1.AddCell(p2);
            pdfTable1.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable1.AddCell(p3);
            pdfTable1.AddCell("");
            pdfTable1.DefaultCell.HorizontalAlignment = Element.ALIGN_RIGHT;
            pdfTable1.AddCell(p4);
            return pdfTable1;

        }

        private PdfPTable CreaTitulos(String xAct, String xUniMedida)
        {
            // Datos necesarios para el titulo
            String tipoInforme = this.comboBoxTipoInforme.SelectedItem.ToString().Substring(2);
            String agrupacion = this.comboBoxAgrupacion.SelectedItem.ToString().Substring(7);

            // Creamos tabla para el título del informe
            PdfPTable pdfTable2 = new PdfPTable(1);
            pdfTable2.SpacingBefore = 5f;
            pdfTable2.WidthPercentage = 100;
            pdfTable2.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            pdfTable2.DefaultCell.VerticalAlignment = Element.ALIGN_CENTER;
            pdfTable2.DefaultCell.BorderWidth = 0;
            pdfTable2.DefaultCell.BorderWidthTop = 1;

            // Se setea título en caso de ser distinto al del combobox
            Chunk c5 = new Chunk("", FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            Chunk c7 = new Chunk("", FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            Chunk c8 = new Chunk("", FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            String titulo = "";
            if (this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("1"))
            {
                c5 = new Chunk(tipoInforme + " (" + xUniMedida + ") x " + agrupacion, FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            }
            else if (this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("2"))
            {
                c5 = new Chunk("Presentación de Proyecciones Contribución " + "(" + xUniMedida + ") x " + agrupacion, FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            }
            else if (this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("3") || this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("4"))
            {
                c5 = new Chunk(tipoInforme + " " + xAct + " x " + agrupacion, FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            }
            else if (this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("5"))
            {
                c5 = new Chunk("Presentación de Planes de Acción " + " " + xAct + " x " + agrupacion, FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            }
            else if (this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("6"))
            {
                c5 = new Chunk("Presentación de Correc. Efect. Comerciales " + " " + xAct + " x " + agrupacion, FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            }
            else if (this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("7"))
            {
                c5 = new Chunk("Ranking de Crecimiento por " + agrupacion, FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
                c7 = new Chunk("Plan Actual v/s Plan Definitivo (" + this.comboBoxVariable.SelectedItem.ToString().Substring(3) + ")", FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
                c8 = new Chunk("Producto ó Grupo de Productos " + this.comboBoxProducto.SelectedItem.ToString(), FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD));
            }
            /*
            // Generamos datos de la tabla de título
            c5 = new Chunk(tipoInforme + " (" + xUniMedida + ") x " + agrupacion, FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            // Validamos si va la unidad de medida o año en el título
            if (this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0,1).Equals("3") || this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("4") || this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("5")) {
                c5 = new Chunk(tipoInforme + " " + xAct + " x " + agrupacion, FontFactory.GetFont("Microsoft Sans Serif", 12, iTextSharp.text.Font.BOLD));
            }
            */
            c5.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);

            Chunk c6 = new Chunk("PROGRAMA OPERATIVO ANUAL " + xAct, FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD));
            c6.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);

            // Validamos si debemos agregar más texto a los títulos según número de informe
            if (this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1).Equals("7"))
            {
                Phrase p5 = new Phrase();
                p5.Add(c5);
                Phrase p7 = new Phrase();
                p7.Add(c7);
                Phrase p8 = new Phrase();
                p8.Add(c8);
                Phrase p6 = new Phrase();
                p6.Add(c6);
                pdfTable2.AddCell(p5);
                pdfTable2.DefaultCell.BorderWidthTop = 0;
                pdfTable2.AddCell(p7);
                pdfTable2.DefaultCell.BorderWidthTop = 0;
                pdfTable2.AddCell(p8);
                
                pdfTable2.DefaultCell.BorderWidthTop = 0;
                pdfTable2.DefaultCell.BorderWidthBottom = 1;
                pdfTable2.DefaultCell.PaddingBottom = 5;
                
                pdfTable2.AddCell(p6);
            }
            else
            {
                Phrase p5 = new Phrase();
                p5.Add(c5);
                Phrase p6 = new Phrase();
                p6.Add(c6);
                pdfTable2.AddCell(p5);
                pdfTable2.DefaultCell.BorderWidthTop = 0;
                pdfTable2.DefaultCell.BorderWidthBottom = 1;
                pdfTable2.DefaultCell.PaddingBottom = 5;
                pdfTable2.AddCell(p6);
            }

            

            return pdfTable2;
        }

        private PdfPTable CreaSubTitulo()
        {
            // Creamos tabla para el subtítulo del informe
            PdfPTable pdfTable3 = new PdfPTable(1);
            pdfTable3.WidthPercentage = 100;
            pdfTable3.DefaultCell.HorizontalAlignment = Element.ALIGN_LEFT;
            pdfTable3.DefaultCell.VerticalAlignment = Element.ALIGN_CENTER;
            pdfTable3.DefaultCell.BorderWidth = 0;
            pdfTable3.DefaultCell.BorderWidthBottom = 1;
            pdfTable3.DefaultCell.PaddingBottom = 5;

            // Validamos que nombre poner en el subtitulo según el tipo de usuario
            String subtitulo = "";
            if (this.comboBoxEntidad.SelectedItem.ToString().Substring(0, 1).Equals("P"))
                subtitulo = "Total Pais";
            else if (this.comboBoxEntidad.SelectedItem.ToString().Substring(0, 1).Equals("R"))
                subtitulo = "Rubro-> " + this.comboBoxEntidad.SelectedItem.ToString();
            else
                subtitulo = this.comboBoxEntidad.SelectedItem.ToString();

            
            // Generamos datos de la tabla de subtitulo
            Chunk c7 = new Chunk(subtitulo, FontFactory.GetFont("Microsoft Sans Serif", 8));
            c7.Font.Color = new iTextSharp.text.BaseColor(0, 0, 0);
            Phrase p7 = new Phrase();
            p7.Add(c7);
            pdfTable3.AddCell(p7);

            return pdfTable3;

        }

        private PdfPTable CreaContenidoInforme1(String xOld, String xAnt, String xAct, String xUniMedida)
        {
            // Creando tabla para guardar DataTable data
            PdfPTable pdfTable = new PdfPTable(7);
            pdfTable.SpacingBefore = 10f;
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 95;
            pdfTable.DefaultCell.BorderWidth = 1f;

            float[] TamColum = new float[] { 1f, 0.6f, 0.6f, 0.6f, 0.6f, 0.6f, 0.6f };
            pdfTable.SetWidths(TamColum);

            // Seteamos la cabecera de la tabla de resultados
            PdfPCell cell1 = new PdfPCell(new Phrase("Producto / Linea", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell1.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.BorderWidthBottom = 0;
            PdfPCell cell2 = new PdfPCell(new Phrase("Venta " + xOld.Substring(2) + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.BorderWidthBottom = 0;
            PdfPCell cell3 = new PdfPCell(new Phrase("Venta " + xAnt.Substring(2) + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell3.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell3.HorizontalAlignment = Element.ALIGN_CENTER;
            cell3.BorderWidthBottom = 0;
            PdfPCell cell4 = new PdfPCell(new Phrase("Poa " + xAct.Substring(2) + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell4.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell4.HorizontalAlignment = Element.ALIGN_CENTER;
            cell4.BorderWidthBottom = 0;
            PdfPCell cell5 = new PdfPCell(new Phrase("Acum.Ene-Sep (%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell5.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell5.HorizontalAlignment = Element.ALIGN_CENTER;
            cell5.BorderWidthBottom = 0;
            PdfPCell cell6 = new PdfPCell(new Phrase("Venta " + xAnt.Substring(2) + "/" + xOld.Substring(2) + "(%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell6.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell6.HorizontalAlignment = Element.ALIGN_CENTER;
            cell6.BorderWidthBottom = 0;
            PdfPCell cell7 = new PdfPCell(new Phrase("Poa " + xAct.Substring(2) + "/" + xAnt.Substring(2) + "(%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell7.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell7.HorizontalAlignment = Element.ALIGN_CENTER;
            cell7.BorderWidthBottom = 0;
            pdfTable.AddCell(cell1);
            pdfTable.AddCell(cell2);
            pdfTable.AddCell(cell3);
            pdfTable.AddCell(cell4);
            pdfTable.AddCell(cell5);
            pdfTable.AddCell(cell6);
            pdfTable.AddCell(cell7);

            /*
            //Agregando los Header de la tabla
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                cell.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
                pdfTable.AddCell(cell);
            }
            */

            // Booleano para validar si la fila contiene un total
            Boolean total = false;

            //Agregando los DataRow
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                total = false;
                foreach (DataGridViewCell cell in row.Cells)
                {

                    // PARA GENERAR BORDES EN LOS TOTALES
                    String texto = ToStringNullSafe(cell.Value);
                    if (texto.ToLower().IndexOf("total") != -1)
                    {
                        cell1.Border = Rectangle.BOTTOM_BORDER | Rectangle.TOP_BORDER;
                        cell1.BorderWidthBottom = 1f;
                        cell1.BorderWidthTop = 1f;
                        total = true;
                    }
                    else
                    {
                        cell1.BorderWidthTop = 0;
                        // REVISAR ****
                    }
                    // Validamos si es fila Total
                    if (total)
                    {
                        cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
                    }
                    else
                    {
                        cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8)));
                        cell1.Border = 0;

                    }

                    cell1.BorderWidthLeft = 0.5f;
                    cell1.BorderWidthRight = 0.5f;

                    // Para alinear el contenido a la derecha
                    if (cell.ColumnIndex != 0)
                    {
                        cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
                    }

                    // REVISAR ****

                    pdfTable.AddCell(cell1);
                }
            }
            return pdfTable;
        }

        private PdfPTable CreaContenidoInforme2(String xOld, String xAnt, String xAct, String xUniMedida)
        {
            // Creando tabla para guardar DataTable data
            PdfPTable pdfTable = new PdfPTable(13);
            pdfTable.SpacingBefore = 10f;
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 98;
            pdfTable.DefaultCell.BorderWidth = 1f;

            float[] TamColum = new float[] { 1f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f, 0.5f };
            pdfTable.SetWidths(TamColum);

            // Seteamos la cabecera de la tabla de resultados
            PdfPCell cell1 = new PdfPCell(new Phrase("Rubros", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell1.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.BorderWidthBottom = 0;
            PdfPCell cell2 = new PdfPCell(new Phrase("Cont. " + xOld.Substring(2) + " (M$)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.BorderWidthBottom = 0;
            PdfPCell cell3 = new PdfPCell(new Phrase("Cont. " + xAnt.Substring(2) + " (M$)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell3.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell3.HorizontalAlignment = Element.ALIGN_CENTER;
            cell3.BorderWidthBottom = 0;
            PdfPCell cell4 = new PdfPCell(new Phrase("Poa " + xAct.Substring(2) + " (M$)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell4.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell4.HorizontalAlignment = Element.ALIGN_CENTER;
            cell4.BorderWidthBottom = 0;
            PdfPCell cell5 = new PdfPCell(new Phrase("Sep. " + xAnt.Substring(2) + "/" + xOld.Substring(2) + " (%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell5.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell5.HorizontalAlignment = Element.ALIGN_CENTER;
            cell5.BorderWidthBottom = 0;
            PdfPCell cell6 = new PdfPCell(new Phrase("Cont. " + xAnt.Substring(2) + "/" + xOld.Substring(2) + " (%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell6.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell6.HorizontalAlignment = Element.ALIGN_CENTER;
            cell6.BorderWidthBottom = 0;
            PdfPCell cell7 = new PdfPCell(new Phrase("Poa " + xAct.Substring(2) + "/" + xAnt.Substring(2) + " (%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell7.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell7.HorizontalAlignment = Element.ALIGN_CENTER;
            cell7.BorderWidthBottom = 0;
            PdfPCell cell8 = new PdfPCell(new Phrase("Mar. " + xOld.Substring(2) + " ($ / Lt)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell8.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell8.HorizontalAlignment = Element.ALIGN_CENTER;
            cell8.BorderWidthBottom = 0;
            PdfPCell cell9 = new PdfPCell(new Phrase("Mar. " + xAnt.Substring(2) + " ($ / Lt)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell9.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell9.HorizontalAlignment = Element.ALIGN_CENTER;
            cell9.BorderWidthBottom = 0;
            PdfPCell cell10 = new PdfPCell(new Phrase("Poa " + xAct.Substring(2) + " ($ / Lt)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell10.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell10.HorizontalAlignment = Element.ALIGN_CENTER;
            cell10.BorderWidthBottom = 0;
            PdfPCell cell11 = new PdfPCell(new Phrase("Sep. " + xAnt.Substring(2) + "/" + xOld.Substring(2) + " (%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell11.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell11.HorizontalAlignment = Element.ALIGN_CENTER;
            cell11.BorderWidthBottom = 0;
            PdfPCell cell12 = new PdfPCell(new Phrase("Mar. " + xAnt.Substring(2) + "/" + xOld.Substring(2) + " (%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell12.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell12.HorizontalAlignment = Element.ALIGN_CENTER;
            cell12.BorderWidthBottom = 0;
            PdfPCell cell13 = new PdfPCell(new Phrase("Poa " + xAct.Substring(2) + "/" + xAnt.Substring(2) + " (%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell13.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell13.HorizontalAlignment = Element.ALIGN_CENTER;
            cell13.BorderWidthBottom = 0;

            pdfTable.AddCell(cell1);
            pdfTable.AddCell(cell2);
            pdfTable.AddCell(cell3);
            pdfTable.AddCell(cell4);
            pdfTable.AddCell(cell5);
            pdfTable.AddCell(cell6);
            pdfTable.AddCell(cell7);
            pdfTable.AddCell(cell8);
            pdfTable.AddCell(cell9);
            pdfTable.AddCell(cell10);
            pdfTable.AddCell(cell11);
            pdfTable.AddCell(cell12);
            pdfTable.AddCell(cell13);

            /*
            //Agregando los Header de la tabla
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                cell.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
                pdfTable.AddCell(cell);
            }
            */

            //Agregando los DataRow
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8)));
                    cell1.Border = 0;
                    cell1.BorderWidthLeft = 0.5f;
                    cell1.BorderWidthRight = 0.5f;
                    // Para alinear el contenido a la derecha
                    if (cell.ColumnIndex != 0)
                    {
                        cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
                    }

                    pdfTable.AddCell(cell1);
                }
            }
            return pdfTable;
        }

        private PdfPTable CreaContenidoInforme3(String xOld, String xAnt, String xAct, String xUniMedida)
        {
            // Creando tabla para guardar DataTable data
            PdfPTable pdfTable = new PdfPTable(7);
            pdfTable.SpacingBefore = 10f;
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 95;
            pdfTable.DefaultCell.BorderWidth = 1f;


            float[] TamColum = new float[] { 1f, 0.6f, 0.6f, 0.6f, 0.6f, 0.6f, 0.6f };
            pdfTable.SetWidths(TamColum);

            // Seteamos la cabecera de la tabla de resultados
            PdfPCell cell1 = new PdfPCell(new Phrase("Producto / Linea", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell1.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.BorderWidthBottom = 0;
            PdfPCell cell2 = new PdfPCell(new Phrase("Plan Teórico " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.BorderWidthBottom = 0;
            PdfPCell cell3 = new PdfPCell(new Phrase("Efectos " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell3.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell3.HorizontalAlignment = Element.ALIGN_CENTER;
            cell3.BorderWidthBottom = 0;
            PdfPCell cell4 = new PdfPCell(new Phrase("Planes Acción " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell4.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell4.HorizontalAlignment = Element.ALIGN_CENTER;
            cell4.BorderWidthBottom = 0;
            PdfPCell cell5 = new PdfPCell(new Phrase("Planes Definit. " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell5.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell5.HorizontalAlignment = Element.ALIGN_CENTER;
            cell5.BorderWidthBottom = 0;
            PdfPCell cell6 = new PdfPCell(new Phrase("(%) Ef.Comer. (EC/PT)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell6.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell6.HorizontalAlignment = Element.ALIGN_CENTER;
            cell6.BorderWidthBottom = 0;
            PdfPCell cell7 = new PdfPCell(new Phrase("(%) P.Acc. (PA/(PT+EC))", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell7.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell7.HorizontalAlignment = Element.ALIGN_CENTER;
            cell7.BorderWidthBottom = 0;
            pdfTable.AddCell(cell1);
            pdfTable.AddCell(cell2);
            pdfTable.AddCell(cell3);
            pdfTable.AddCell(cell4);
            pdfTable.AddCell(cell5);
            pdfTable.AddCell(cell6);
            pdfTable.AddCell(cell7);

            // Booleano para validar si la fila contiene un total
            Boolean total = false;

            //Agregando los DataRow
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                total = false;
                foreach (DataGridViewCell cell in row.Cells)
                {

                    // PARA GENERAR BORDES EN LOS TOTALES
                    String texto = ToStringNullSafe(cell.Value);
                    if (texto.ToLower().IndexOf("total") != -1)
                    {
                        cell1.Border = Rectangle.BOTTOM_BORDER | Rectangle.TOP_BORDER;
                        cell1.BorderWidthBottom = 1f;
                        cell1.BorderWidthTop = 1f;
                        total = true;
                    }
                    else
                    {
                        cell1.BorderWidthTop = 0;
                        // REVISAR ****
                    }
                    // Validamos si es fila Total
                    if (total)
                    {
                        cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
                    }
                    else {
                        cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8)));
                        cell1.Border = 0;

                    }

                    cell1.BorderWidthLeft = 0.5f;
                    cell1.BorderWidthRight = 0.5f;

                    // Para alinear el contenido a la derecha
                    if (cell.ColumnIndex != 0)
                    {
                        cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
                    }

                    // REVISAR ****
              
                    pdfTable.AddCell(cell1);
                }
            }
            return pdfTable;
        }

        private PdfPTable CreaContenidoInforme4(String xAct, String xUniMedida)
        {
            // Creando tabla para guardar DataTable data
            PdfPTable pdfTable = new PdfPTable(14);
            pdfTable.SpacingBefore = 10f;
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 100;
            pdfTable.DefaultCell.BorderWidth = 1f;


            float[] TamColum = new float[] { 1f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f, 0.4f };
            pdfTable.SetWidths(TamColum);

            // Seteamos la cabecera de la tabla de resultados
            PdfPCell cell1 = new PdfPCell(new Phrase("Producto / Linea", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell1.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.BorderWidthBottom = 0;
            PdfPCell cell2 = new PdfPCell(new Phrase("Ene." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.BorderWidthBottom = 0;
            PdfPCell cell3 = new PdfPCell(new Phrase("Feb." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell3.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell3.HorizontalAlignment = Element.ALIGN_CENTER;
            cell3.BorderWidthBottom = 0;
            PdfPCell cell4 = new PdfPCell(new Phrase("Mar." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell4.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell4.HorizontalAlignment = Element.ALIGN_CENTER;
            cell4.BorderWidthBottom = 0;
            PdfPCell cell5 = new PdfPCell(new Phrase("Abr." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell5.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell5.HorizontalAlignment = Element.ALIGN_CENTER;
            cell5.BorderWidthBottom = 0;
            PdfPCell cell6 = new PdfPCell(new Phrase("May." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell6.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell6.HorizontalAlignment = Element.ALIGN_CENTER;
            cell6.BorderWidthBottom = 0;
            PdfPCell cell7 = new PdfPCell(new Phrase("Jun." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell7.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell7.HorizontalAlignment = Element.ALIGN_CENTER;
            cell7.BorderWidthBottom = 0;
            PdfPCell cell8 = new PdfPCell(new Phrase("Jul." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell8.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell8.HorizontalAlignment = Element.ALIGN_CENTER;
            cell8.BorderWidthBottom = 0;
            PdfPCell cell9 = new PdfPCell(new Phrase("Ago." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell9.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell9.HorizontalAlignment = Element.ALIGN_CENTER;
            cell9.BorderWidthBottom = 0;
            PdfPCell cell10 = new PdfPCell(new Phrase("Sep." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell10.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell10.HorizontalAlignment = Element.ALIGN_CENTER;
            cell10.BorderWidthBottom = 0;
            PdfPCell cell11 = new PdfPCell(new Phrase("Oct." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell11.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell11.HorizontalAlignment = Element.ALIGN_CENTER;
            cell11.BorderWidthBottom = 0;
            PdfPCell cell12 = new PdfPCell(new Phrase("Nov." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell12.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell12.HorizontalAlignment = Element.ALIGN_CENTER;
            cell12.BorderWidthBottom = 0;
            PdfPCell cell13 = new PdfPCell(new Phrase("Dic." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell13.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell13.HorizontalAlignment = Element.ALIGN_CENTER;
            cell13.BorderWidthBottom = 0;
            PdfPCell cell14 = new PdfPCell(new Phrase("Total." + xAct + " (" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell14.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell14.HorizontalAlignment = Element.ALIGN_CENTER;
            cell14.BorderWidthBottom = 0;
            pdfTable.AddCell(cell1);
            pdfTable.AddCell(cell2);
            pdfTable.AddCell(cell3);
            pdfTable.AddCell(cell4);
            pdfTable.AddCell(cell5);
            pdfTable.AddCell(cell6);
            pdfTable.AddCell(cell7);
            pdfTable.AddCell(cell8);
            pdfTable.AddCell(cell9);
            pdfTable.AddCell(cell10);
            pdfTable.AddCell(cell11);
            pdfTable.AddCell(cell12);
            pdfTable.AddCell(cell13);
            pdfTable.AddCell(cell14);

            // Booleano para validar si la fila contiene un total
            Boolean total = false;

            //Agregando los DataRow
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                total = false;
                foreach (DataGridViewCell cell in row.Cells)
                {

                    // PARA GENERAR BORDES EN LOS TOTALES
                    String texto = ToStringNullSafe(cell.Value);
                    if (texto.ToLower().IndexOf("total") != -1)
                    {
                        cell1.Border = Rectangle.BOTTOM_BORDER | Rectangle.TOP_BORDER;
                        cell1.BorderWidthBottom = 1f;
                        cell1.BorderWidthTop = 1f;
                        total = true;
                    }
                    else
                    {
                        cell1.BorderWidthTop = 0;
                        // REVISAR ****
                    }
                    // Validamos si es fila Total
                    if (total)
                    {
                        cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
                    }
                    else
                    {
                        cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8)));
                        cell1.Border = 0;

                    }

                    cell1.BorderWidthLeft = 0.5f;
                    cell1.BorderWidthRight = 0.5f;

                    // Para alinear el contenido a la derecha
                    if (cell.ColumnIndex != 0)
                    {
                        cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
                    }

                    // REVISAR ****

                    pdfTable.AddCell(cell1);
                }
            }
            return pdfTable;
        }

        private PdfPTable CreaContenidoInforme5(String xUniMedida)
        {
            // Creando tabla para guardar DataTable data
            PdfPTable pdfTable = new PdfPTable(7);
            pdfTable.SpacingBefore = 10f;
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 95;
            pdfTable.DefaultCell.BorderWidth = 1f;


            float[] TamColum = new float[] { 1f, 0.6f, 0.6f, 0.6f, 0.6f, 0.6f, 0.6f };
            pdfTable.SetWidths(TamColum);

            // Seteamos la cabecera de la tabla de resultados
            PdfPCell cell1 = new PdfPCell(new Phrase("Producto / Linea", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell1.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.BorderWidthBottom = 0;
            PdfPCell cell2 = new PdfPCell(new Phrase("Precio " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.BorderWidthBottom = 0;
            PdfPCell cell3 = new PdfPCell(new Phrase("Prod/Ser " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell3.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell3.HorizontalAlignment = Element.ALIGN_CENTER;
            cell3.BorderWidthBottom = 0;
            PdfPCell cell4 = new PdfPCell(new Phrase("F.Venta " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell4.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell4.HorizontalAlignment = Element.ALIGN_CENTER;
            cell4.BorderWidthBottom = 0;
            PdfPCell cell5 = new PdfPCell(new Phrase("Prom/Pub " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell5.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell5.HorizontalAlignment = Element.ALIGN_CENTER;
            cell5.BorderWidthBottom = 0;
            PdfPCell cell6 = new PdfPCell(new Phrase("Operac. " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell6.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell6.HorizontalAlignment = Element.ALIGN_CENTER;
            cell6.BorderWidthBottom = 0;
            PdfPCell cell7 = new PdfPCell(new Phrase("Total " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell7.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell7.HorizontalAlignment = Element.ALIGN_CENTER;
            cell7.BorderWidthBottom = 0;
            pdfTable.AddCell(cell1);
            pdfTable.AddCell(cell2);
            pdfTable.AddCell(cell3);
            pdfTable.AddCell(cell4);
            pdfTable.AddCell(cell5);
            pdfTable.AddCell(cell6);
            pdfTable.AddCell(cell7);

            // Booleano para validar si la fila contiene un total
            Boolean total = false;

            //Agregando los DataRow
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                total = false;
                foreach (DataGridViewCell cell in row.Cells)
                {

                    // PARA GENERAR BORDES EN LOS TOTALES
                    String texto = ToStringNullSafe(cell.Value);
                    if (texto.ToLower().IndexOf("total") != -1)
                    {
                        cell1.Border = Rectangle.BOTTOM_BORDER | Rectangle.TOP_BORDER;
                        cell1.BorderWidthBottom = 1f;
                        cell1.BorderWidthTop = 1f;
                        total = true;
                    }
                    else
                    {
                        cell1.BorderWidthTop = 0;
                        // REVISAR ****
                    }
                    // Validamos si es fila Total
                    if (total)
                    {
                        cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
                    }
                    else
                    {
                        cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 8)));
                        cell1.Border = 0;

                    }

                    cell1.BorderWidthLeft = 0.5f;
                    cell1.BorderWidthRight = 0.5f;

                    // Para alinear el contenido a la derecha
                    if (cell.ColumnIndex != 0)
                    {
                        cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
                    }

                    // REVISAR ****

                    pdfTable.AddCell(cell1);
                }
            }
            return pdfTable;
        }

        private PdfPTable CreaContenidoInforme7(String xUniMedida)
        {
            // Creando tabla para guardar DataTable data
            PdfPTable pdfTable = new PdfPTable(6);
            pdfTable.SpacingBefore = 10f;
            pdfTable.DefaultCell.Padding = 3;
            pdfTable.WidthPercentage = 95;
            pdfTable.DefaultCell.BorderWidth = 1f;


            float[] TamColum = new float[] { 0.44f, 1f, 0.3f, 0.3f, 0.3f, 0.23f };
            pdfTable.SetWidths(TamColum);

            // Seteamos la cabecera de la tabla de resultados
            PdfPCell cell1 = new PdfPCell(new Phrase("Cód Cliente", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell1.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell1.HorizontalAlignment = Element.ALIGN_CENTER;
            cell1.BorderWidthBottom = 0;
            PdfPCell cell2 = new PdfPCell(new Phrase("Razón Social", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell2.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell2.HorizontalAlignment = Element.ALIGN_CENTER;
            cell2.BorderWidthBottom = 0;
            PdfPCell cell3 = new PdfPCell(new Phrase("Plan Actual " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell3.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell3.HorizontalAlignment = Element.ALIGN_CENTER;
            cell3.BorderWidthBottom = 0;
            PdfPCell cell4 = new PdfPCell(new Phrase("Plan Definitivo " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell4.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell4.HorizontalAlignment = Element.ALIGN_CENTER;
            cell4.BorderWidthBottom = 0;
            PdfPCell cell5 = new PdfPCell(new Phrase("Diferencia " + "(" + xUniMedida + ")", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell5.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell5.HorizontalAlignment = Element.ALIGN_CENTER;
            cell5.BorderWidthBottom = 0;
            PdfPCell cell6 = new PdfPCell(new Phrase("Crecimiento (%)", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
            cell6.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
            cell6.HorizontalAlignment = Element.ALIGN_CENTER;
            cell6.BorderWidthBottom = 0;
            pdfTable.AddCell(cell1);
            pdfTable.AddCell(cell2);
            pdfTable.AddCell(cell3);
            pdfTable.AddCell(cell4);
            pdfTable.AddCell(cell5);
            pdfTable.AddCell(cell6);

            // Para contabilizar el número de filas
            int filasAux = 0;
            int filasTotal = 0;

            //Agregando los DataRow
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell1 = new PdfPCell(new Phrase(cell.Value?.ToString() ?? "", FontFactory.GetFont("Microsoft Sans Serif", 7)));
                    cell1.Border = 0;
                    cell1.BorderWidthLeft = 0.5f;
                    cell1.BorderWidthRight = 0.5f;
                    // Para alinear el contenido a la derecha
                    if (cell.ColumnIndex != 0 && cell.ColumnIndex != 1)
                    {
                        cell1.HorizontalAlignment = Element.ALIGN_RIGHT;
                    }

                    // Validamos si es la última fila para pintar el borde inferior
                    if (filasTotal >= dataGridView1.RowCount-1)
                    {
                        cell1.Border = Rectangle.BOTTOM_BORDER;
                        cell1.BorderWidthBottom = 1f;
                    }
                    pdfTable.AddCell(cell1);
                    
                }
                // Validamos cuando insertar los títulos
                filasAux++;
                filasTotal++;
                if (filasAux == 54) {
                    // Restamos 8 tomando en cuenta que desde la segunda página no hay Header
                    filasAux = -11;
                    cell1 = new PdfPCell(new Phrase("Cód Cliente", FontFactory.GetFont("Microsoft Sans Serif", 8, iTextSharp.text.Font.BOLD)));
                    cell1.BackgroundColor = new iTextSharp.text.BaseColor(240, 240, 240);
                    cell1.HorizontalAlignment = Element.ALIGN_CENTER;
                    cell1.BorderWidthBottom = 0;
                    pdfTable.AddCell(cell1);
                    pdfTable.AddCell(cell2);
                    pdfTable.AddCell(cell3);
                    pdfTable.AddCell(cell4);
                    pdfTable.AddCell(cell5);
                    pdfTable.AddCell(cell6);
                }

            }
            return pdfTable;
        }

        // Método para validar si texto en objeto es nulo antes de asignarlo
        private string ToStringNullSafe(object value)
        {
            return (value ?? string.Empty).ToString();
        }

        private void ArmadoPDF(String tipoInforme, String canal, String xAct, String xUniMedida, String xOld, String xAnt)
        {
            // ARMADO PDF
            // Directorio de destino
            string folderPath = "C:\\PDFs\\";

            //Exportando a PDF
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (FileStream stream = new FileStream(folderPath + "Informe " + tipoInforme + ".pdf", FileMode.Create))
            {

                // Seteamos el tipo de Hoja dependiendo del informe seleccionado
                // Tamaño Carta
                Document pdfDoc = new Document(PageSize.LETTER, 23f, 23f, 23f, 30f);
                if (!tipoInforme.Equals("1") && !tipoInforme.Equals("3") && !tipoInforme.Equals("5") && !tipoInforme.Equals("7") && !tipoInforme.Equals("8") && !tipoInforme.Equals("9"))
                {
                    // Para estilo Landscape
                    pdfDoc = new Document(PageSize.LETTER.Rotate(), 23f, 23f, 23f, 30f);
                }
        
                // Generamos la instancia
                PdfWriter writer = PdfWriter.GetInstance(pdfDoc, stream);

                // Para la escritura de Footer en cada página que se crea
                writer.PageEvent = new ITextEvents();

                // Pintamos el PDF
                pdfDoc.Open();
                pdfDoc.AddTitle(this.comboBoxTipoInforme.SelectedItem.ToString().Substring(2));
                PdfPTable header = CreaHeader(canal);
                PdfPTable titulos = CreaTitulos(xAct, xUniMedida);
                PdfPTable subtitulo = CreaSubTitulo();
                pdfDoc.Add(header);
                pdfDoc.Add(titulos);
                pdfDoc.Add(subtitulo);

                // Pintamos el contenido según el informe que se seleccionó
                if (tipoInforme.Equals("1"))
                {
                    PdfPTable contenido = CreaContenidoInforme1(xOld, xAnt, xAct, xUniMedida);
                    pdfDoc.Add(contenido);
                }
                else if (tipoInforme.Equals("2"))
                {
                    PdfPTable contenido = CreaContenidoInforme2(xOld, xAnt, xAct, xUniMedida);
                    pdfDoc.Add(contenido);
                }
                //OJO con el contenido
                else if (tipoInforme.Equals("3"))
                {
                    PdfPTable contenido = CreaContenidoInforme3(xOld, xAnt, xAct, xUniMedida);
                    pdfDoc.Add(contenido);
                }
                else if (tipoInforme.Equals("4"))
                {
                    PdfPTable contenido = CreaContenidoInforme4(xAct, xUniMedida);
                    pdfDoc.Add(contenido);
                }
                else if (tipoInforme.Equals("5"))
                {
                    PdfPTable contenido = CreaContenidoInforme5(xUniMedida);
                    pdfDoc.Add(contenido);
                }
                else if (tipoInforme.Equals("7"))
                {
                    PdfPTable contenido = CreaContenidoInforme7(xUniMedida);
                    pdfDoc.Add(contenido);
                }
                // Cerramos
                pdfDoc.Close();
                stream.Close();
            }
            // Iniciamos aplicación que abre PDF por defecto con el informe generado
            Process.Start(folderPath + "Informe " + tipoInforme + ".pdf");

        }
        #endregion

        #region Botones
        // Botón imprimir
        private void button1_Click(object sender, EventArgs e)
        {
            // Implementar código para validar los combobox y según eso ir al informe que corresponda

            // Parámetros de los combobox
            String variable = this.comboBoxVariable.SelectedItem.ToString().Substring(0, 1);
            String tipoInforme = this.comboBoxTipoInforme.SelectedItem.ToString().Substring(0, 1);
            String agrupacion = this.comboBoxAgrupacion.SelectedItem.ToString().Substring(0, 1);
            String producto = this.comboBoxProducto.SelectedItem.ToString().Substring(0, 1);
            String ordenamiento = this.comboBoxOrden.SelectedItem.ToString().Substring(0, 1);

            // Parámetros de entrada
            String xOld = "2016";
            String xAnt = "2017";
            String xAct = "2018";
            String xUniMedida = "M3";
            // Esto se debe parametrizar
            //String tipoUsuario = "P";
            //String tipoUsuario = "R";
            //String tipoUsuario = "O";           
            //String tipoUsuario = "Z";
            //String canal = "Concesionarios";


            // Llamamos a método que obtiene datos según número de informe
            Console.WriteLine("Informe: " + tipoInforme);

            // Obtención de datos desde BD para pintar en Informe
            //DatosBD(tipoInforme);
            DataInformesBD(tipoInforme);

            // Armado de archivo PDF con los datos del informe correspondiente
            ArmadoPDF(tipoInforme, this.canal, xAct, xUniMedida, xOld, xAnt);

        }

        // Botón Salir
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion
    }
}