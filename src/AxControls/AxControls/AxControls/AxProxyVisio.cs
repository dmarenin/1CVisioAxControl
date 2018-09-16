using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Drawing;
using System.Collections.Generic;

using VisOcx = AxMicrosoft.Office.Interop.VisOcx;
using Visio = Microsoft.Office.Interop.Visio;

using GMap.NET;
using GMap.NET.WindowsForms;
using GMap.NET.ObjectModel;
using GMap.NET.MapProviders;
using GMap.NET.WindowsForms.Markers;
using GMap.NET.WindowsForms.ToolTips;

namespace AxControls
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F78344A8-DEDF-49C4-A04C-17A1C8C6294D")]
    [ProgId("AxControls.AxProxyVisio")]
    [ComDefaultInterface(typeof(IAxProxyVisio))]
    [ComSourceInterfaces(typeof(IAxProxyVisioEvents))]
    public class AxProxyVisio : UserControl, IAxProxyVisio, IObjectSafety
    {
        #region IObjectSafety
        public enum ObjectSafetyOptions
        {
            INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001,
            INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002,
            INTERFACE_USES_DISPEX = 0x00000004,
            INTERFACE_USES_SECURITY_MANAGER = 0x00000008
        };

        public int GetInterfaceSafetyOptions(ref Guid riid, out int pdwSupportedOptions, out int pdwEnabledOptions)
        {
            ObjectSafetyOptions m_options = ObjectSafetyOptions.INTERFACESAFE_FOR_UNTRUSTED_CALLER
                | ObjectSafetyOptions.INTERFACESAFE_FOR_UNTRUSTED_DATA
                | ObjectSafetyOptions.INTERFACE_USES_DISPEX
                | ObjectSafetyOptions.INTERFACE_USES_SECURITY_MANAGER;

            pdwSupportedOptions = (int)m_options;
            pdwEnabledOptions = (int)m_options;
            return 0;
        }
        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            return 0;
        }

        #endregion

        #region AxControls

        private System.ComponentModel.IContainer components = null;

        public string pathFile;
        string IAxProxyVisio.PathFile
        {
            get { return pathFile; }
            set { pathFile = value; }
        }

        public string owner1c;
        string IAxProxyVisio.Owner1c
        {
            get { return owner1c; }
            set { owner1c = value; }
        }

        public VisOcx.AxDrawingControl axDrawingControl1;
        object IAxProxyVisio.AxDrawingControl
        {
            get { return axDrawingControl1; }
        }

        public string Version
        {
            get
            {
                return this.GetType().Assembly.GetName().Version.ToString();
            }
        }

        public PictureBox pictureBox1;
        PictureBox IAxProxyVisio.PictureLoading
        {
            get { return pictureBox1; }
        }

        public Visio.Master GetMaster(Visio.Document currentDoc, string type)
        {
            Visio.Master master = currentDoc.Masters[type];
            return master;
        }

        public Visio.Page GetActivePage(string pageHeight, string pageWidth)
        {
            Visio.Document doc = axDrawingControl1.Document;
            Visio.Page page = doc.Pages[1];
            page.PageSheet.Cells["PageHeight"].FormulaU = pageHeight;
            page.PageSheet.Cells["PageWidth"].FormulaU = pageWidth;
            page.AutoSize = false;
            doc.DiagramServicesEnabled = 0;
            return page;
        }

        public Visio.Application GetApplication()
        {
            return axDrawingControl1.Document.Application;
        }

        public string GetValueProperty(Visio.Shape shape, string propertyName)
        {
            Visio.Cell cell = shape.Cells[propertyName];
            try
            {
                return cell.Formula;
            }
            catch
            {
                return null;
            }
        }

        public bool SetValueProperty(Visio.Shape shape, string propertyName, string propertyValue)
        {
            try
            {
                Visio.Cell cell = shape.Cells[propertyName];
                cell.Formula = propertyValue;
                return true;
            }
            catch
            {
                return false;
            }
        }
        public void SetValuesProperties(dynamic structShape)
        {
            foreach (dynamic str in structShape.МассивСвойств)
            {
                if (structShape.НеИспользоватьСвойства != null)
                {
                    if (structShape.НеИспользоватьСвойства.Find(str.Key) != null)
                    {
                        continue;
                    }
                }
                SetValueProperty(structShape.COMОбъект, str.Key, str.Value);
            }
        }

        public Visio.Cell GetCell(Visio.Shape shape, string propertyName)
        {
            try
            {
                Visio.Cell cell = shape.Cells[propertyName];
                return cell;
            }
            catch
            {
                return null;
            }

        }

        public void Clean()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForFullGCComplete();
            GC.Collect();
        }
        public void Undo()
        {
            Visio.Application app = axDrawingControl1.Window.Application;
            try
            {
                app.DoCmd((short)Visio.VisUICmds.visCmdEditUndo);
            }
            catch
            {
            }
        }
        public void Redo()
        {
            Visio.Application app = axDrawingControl1.Window.Application;
            try
            {
                app.DoCmd((short)Visio.VisUICmds.visCmdEditRedo);
            }
            catch
            {
            }
        }
        public void SetEventDblClick(Visio.Shape shape)
        {
            shape.Cells["EventDblClick"].Formula = "=QUEUEMARKEREVENT(\"/cmd=DoubleClick /source=" + shape.Data1 + "\")";
        }
        public void AddCustomProperty(Visio.Shape shape, string propertyName, string propertyValue)
        {
            short customProps = (short)Visio.VisSectionIndices.visSectionProp;
            short rowNumber = shape.AddRow(customProps, (short)Visio.VisRowIndices.visRowLast, (short)Visio.VisRowTags.visTagDefault);
            shape.CellsSRC[customProps, rowNumber, (short)Visio.VisCellIndices.visCustPropsLabel].FormulaU = "\"" + propertyName + "\"";
            shape.CellsSRC[customProps, rowNumber, (short)Visio.VisCellIndices.visCustPropsValue].FormulaU = "\"" + propertyValue + "\"";
        }
        public void SendDocument()
        {
            Visio.Document doc = axDrawingControl1.Document;
            doc.Application.DoCmd((short)Visio.VisUICmds.visCmdSendAsMail);
        }
        public void ActivateShape(Visio.Shape shape)
        {
            Visio.Document doc = axDrawingControl1.Document;
            Visio.Window activeWindow = doc.Application.ActiveWindow;
            activeWindow.Zoom = 0.85;
            activeWindow.CenterViewOnShape(shape, Visio.VisCenterViewFlags.visCenterViewSelectShape);
            this.Show();
        }
        public void SaveDocument()
        {
            SaveFileDialog dlgSaveDiagram = new SaveFileDialog();
            dlgSaveDiagram.Filter = "Visio Diagrams|*.vsd | All files(*.*) | *.*";
            Visio.Document doc = axDrawingControl1.Document;
            if (dlgSaveDiagram.ShowDialog() == DialogResult.OK)
            {
                string FileName = dlgSaveDiagram.FileName;
                try
                {
                    doc.SaveAsEx(Path.GetFullPath(FileName), (short)Visio.VisOpenSaveArgs.visSaveAsWS);

                    //this.axDrawingControl1.Document.SaveAs(Path.GetFullPath(FileName));
                }
                catch
                {
                }
            }
        }
        public void PrintDocument()
        {
            Visio.Document doc = axDrawingControl1.Document;
            Visio.Selection select = doc.Application.ActiveWindow.Selection;

            System.Windows.Forms.PrintDialog printDialog = new System.Windows.Forms.PrintDialog();
            if (printDialog.ShowDialog() == DialogResult.OK)
            {
                var PrinterName = printDialog.PrinterSettings.PrinterName;

                doc.PrintOut(Visio.VisPrintOutRange.visPrintCurrentView, 1, -1, true, PrinterName);
            }
        }
        public void Message(string mes = "", string cap = "Ошибка")
        {
            MessageBox.Show(mes, cap);
        }

        //Старая версия
        //public void ConnectShapes(object shape1, object shape2, Visio.Shape connector)
        //{
        //    if (shape1 == null)
        //    {
        //        return;
        //    }
        //    Visio.Cell beginXCell = connector.get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXForm1D, (short)Visio.VisCellIndices.vis1DBeginX);
        //    beginXCell.GlueTo(((Visio.Shape)shape1).get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX));
        //    if (shape2 != null)
        //    {
        //        Visio.Cell endXCell = connector.get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXForm1D, (short)Visio.VisCellIndices.vis1DEndX);
        //        endXCell.GlueTo(((Visio.Shape)shape2).get_CellsSRC((short)Visio.VisSectionIndices.visSectionObject, (short)Visio.VisRowIndices.visRowXFormOut, (short)Visio.VisCellIndices.visXFormPinX));
        //    }
        //}
        public void ConnectShapes(object shape1, object shape2, Visio.Shape connector)
        {
            if (shape1 == null)
            {
                return;
            }

            Visio.Cell beginXCell = connector.get_CellsSRC(1, 4, 0);
            beginXCell.GlueTo(((Visio.Shape)shape1).get_CellsSRC(1, 1, 0));

            if (shape2 != null)
            {
                Visio.Cell endXCell = connector.get_CellsSRC(1, 4, 2);
                endXCell.GlueTo(((Visio.Shape)shape2).get_CellsSRC(1, 1, 0));
            }
        }

        public AxProxyVisio()
        {
            InitializeComponent();
            InitializeMap();

            ActiveControl = axDrawingControl1;

            pictureBox1.Visible = false;
            mainMap.Visible = false;
            mainMap.Enabled = false;

            provider_YandexHybrid = YandexHybridMapProvider.Instance;
            provider_Yandex = YandexMapProvider.Instance;
            provider_YandexSatellite = YandexSatelliteMapProvider.Instance;

            provider_GoogleSatellite = GoogleSatelliteMapProvider.Instance;
            provider_GoogleHybrid = GoogleHybridMapProvider.Instance;
            provider_GoogleTerrain = GoogleTerrainMapProvider.Instance;
            provider_Google = GoogleMapProvider.Instance;

            provider_Empty = EmptyProvider.Instance;

            provider_OpenStreetMap = OpenStreetMapProvider.Instance;
        }

        public void InitializeComponent()
        {
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.mainMap = new AxControls.Map();
            this.axDrawingControl1 = new AxMicrosoft.Office.Interop.VisOcx.AxDrawingControl();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.axDrawingControl1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.White;
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Image = global::AxControls.Properties.Resources.ДлительнаяОперация48;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(907, 589);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // mainMap
            // 
            this.mainMap.BackColor = System.Drawing.SystemColors.Control;
            this.mainMap.Bearing = 0F;
            this.mainMap.CanDragMap = true;
            this.mainMap.Cursor = System.Windows.Forms.Cursors.Default;
            this.mainMap.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainMap.EmptyTileColor = System.Drawing.Color.Navy;
            this.mainMap.GrayScaleMode = false;
            this.mainMap.HelperLineOption = HelperLineOptions.DontShow;
            this.mainMap.LevelsKeepInMemmory = 5;
            this.mainMap.Location = new System.Drawing.Point(0, 0);
            this.mainMap.MarkersEnabled = true;
            this.mainMap.MaxZoom = 17;
            this.mainMap.MinZoom = 2;
            this.mainMap.MouseWheelZoomEnabled = true;
            this.mainMap.MouseWheelZoomType = MouseWheelZoomType.MousePositionAndCenter;
            this.mainMap.Name = "mainMap";
            this.mainMap.NegativeMode = false;
            this.mainMap.PolygonsEnabled = true;
            this.mainMap.RetryLoadTile = 0;
            this.mainMap.RoutesEnabled = true;
            this.mainMap.ScaleMode = ScaleModes.Integer;
            this.mainMap.SelectedAreaFillColor = System.Drawing.Color.FromArgb(((int)(((byte)(33)))), ((int)(((byte)(65)))), ((int)(((byte)(105)))), ((int)(((byte)(225)))));
            this.mainMap.ShowTileGridLines = false;
            this.mainMap.Size = new System.Drawing.Size(907, 589);
            this.mainMap.TabIndex = 2;
            this.mainMap.Zoom = 0D;
            // 
            // axDrawingControl1
            // 
            this.axDrawingControl1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(240)))), ((int)(((byte)(240)))), ((int)(((byte)(240)))));
            this.axDrawingControl1.CausesValidation = false;
            this.axDrawingControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.axDrawingControl1.Enabled = true;
            this.axDrawingControl1.Location = new System.Drawing.Point(0, 0);
            this.axDrawingControl1.Name = "axDrawingControl1";
            this.axDrawingControl1.Size = new System.Drawing.Size(907, 589);
            this.axDrawingControl1.TabIndex = 0;
            // 
            // AxProxyVisio
            // 
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.mainMap);
            this.Controls.Add(this.axDrawingControl1);
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            this.Name = "AxProxyVisio";
            this.Size = new System.Drawing.Size(907, 589);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.axDrawingControl1)).EndInit();
            this.ResumeLayout(false);

        }

        protected override void Dispose(bool disposing)
        {
            Clean();

            this.Controls.Clear();

            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #endregion

        #region GMap

        public Map mainMap;
        Map IAxProxyVisio.MainMap
        {
            get { return mainMap; }
        }

        public readonly GMapProvider provider_YandexHybrid;
        string IAxProxyVisio.id_YandexHybrid
        {
            get { return provider_YandexHybrid.Id.ToString(); }
        }

        public readonly GMapProvider provider_Yandex;
        string IAxProxyVisio.id_Yandex
        {
            get { return provider_Yandex.Id.ToString(); }
        }

        public readonly GMapProvider provider_YandexSatellite;
        string IAxProxyVisio.id_YandexSatellite
        {
            get { return provider_YandexSatellite.Id.ToString(); }
        }

        public readonly GMapProvider provider_GoogleSatellite;
        string IAxProxyVisio.id_GoogleSatellite
        {
            get { return provider_GoogleSatellite.Id.ToString(); }
        }

        public readonly GMapProvider provider_GoogleHybrid;
        string IAxProxyVisio.id_GoogleHybrid
        {
            get { return provider_GoogleHybrid.Id.ToString(); }
        }

        public readonly GMapProvider provider_GoogleTerrain;
        string IAxProxyVisio.id_GoogleTerrain
        {
            get { return provider_GoogleTerrain.Id.ToString(); }
        }

        public readonly GMapProvider provider_Empty;
        string IAxProxyVisio.id_Empty
        {
            get { return provider_Empty.Id.ToString(); }
        }

        public readonly GMapProvider provider_OpenStreetMap;
        string IAxProxyVisio.id_OpenStreetMap
        {
            get { return provider_OpenStreetMap.Id.ToString(); }
        }

        public readonly GMapProvider provider_Google;
        string IAxProxyVisio.id_Google
        {
            get { return provider_Google.Id.ToString(); }
        }

        bool IAxProxyVisio.MapVisible
        {
            get { return mainMap.Visible; }
            set { mainMap.Visible = value; }
        }

        bool IAxProxyVisio.MapEnabled
        {
            get { return mainMap.Enabled; }
            set { mainMap.Enabled = value; }
        }

        public delegate void MapEvent_Delegate(string nameEvent, string args);
        private event MapEvent_Delegate MapEvent;

        public void ChangeMapType(string id)
        {
            mainMap.MapProvider = GMapProviders.TryGetProvider(new Guid(id));
        }

        // layers
        readonly GMapOverlay top = new GMapOverlay();

        //internal readonly GMapOverlay objects = new GMapOverlay("objects");
        public GMapOverlay objects = new GMapOverlay("objects");
        GMapOverlay IAxProxyVisio.MarkerObjects
        {
            get { return objects; }
        }

        public GMapOverlay routes = new GMapOverlay("routes");
        GMapOverlay IAxProxyVisio.MarekerRoutes
        {
            get { return routes; }
        }

        public GMapOverlay polygons = new GMapOverlay("polygons");
        GMapOverlay IAxProxyVisio.MarekerPolygons
        {
            get { return polygons; }
        }

        public void AddMarkers(object overlay, object val)
        {
            ((GMapOverlay)overlay).Markers.Add((GMapMarker)val);
        }

        // marker
        GMapMarker currentMarker;

        // polygons
        GMapPolygon polygon;

        // etc
        readonly Random rnd = new Random();

        GMapMarkerRect CurentRectMarker = null;
        string mobileGpsLog = string.Empty;
        bool isMouseDown = false;
        PointLatLng start;
        PointLatLng end;
        PointLatLng lastPosition;
        int lastZoom;
        private TrackBar trackBar1;

        GMapRoute currentRoute = null;
        GMapPolygon currentPolygon = null;

        public void InitializeMap()
        {
            //mainMap.MapProvider = GMapProviders.OpenStreetMap;
            mainMap.MapProvider = GMapProviders.GoogleMap;

            mainMap.HelperLineOption = HelperLineOptions.ShowOnModifierKey;

            mainMap.Position = new PointLatLng(54.6961334816182, 25.2985095977783);
            mainMap.MinZoom = 0;
            mainMap.MaxZoom = 24;
            mainMap.Zoom = 9;

            mainMap.OnPositionChanged += new PositionChanged(MainMap_OnPositionChanged);

            mainMap.OnTileLoadStart += new TileLoadStart(MainMap_OnTileLoadStart);
            mainMap.OnTileLoadComplete += new TileLoadComplete(MainMap_OnTileLoadComplete);

            mainMap.OnMapZoomChanged += new MapZoomChanged(MainMap_OnMapZoomChanged);
            mainMap.OnMapTypeChanged += new MapTypeChanged(MainMap_OnMapTypeChanged);

            mainMap.OnMarkerClick += new MarkerClick(MainMap_OnMarkerClick);
            mainMap.OnMarkerEnter += new MarkerEnter(MainMap_OnMarkerEnter);
            mainMap.OnMarkerLeave += new MarkerLeave(MainMap_OnMarkerLeave);

            mainMap.OnPolygonEnter += new PolygonEnter(MainMap_OnPolygonEnter);
            mainMap.OnPolygonLeave += new PolygonLeave(MainMap_OnPolygonLeave);

            mainMap.OnRouteEnter += new RouteEnter(MainMap_OnRouteEnter);
            mainMap.OnRouteLeave += new RouteLeave(MainMap_OnRouteLeave);

            mainMap.Manager.OnTileCacheComplete += new TileCacheComplete(OnTileCacheComplete);
            mainMap.Manager.OnTileCacheStart += new TileCacheStart(OnTileCacheStart);
            mainMap.Manager.OnTileCacheProgress += new TileCacheProgress(OnTileCacheProgress);

            mainMap.MouseMove += new MouseEventHandler(MainMap_MouseMove);
            mainMap.MouseDown += new MouseEventHandler(MainMap_MouseDown);
            mainMap.MouseUp += new MouseEventHandler(MainMap_MouseUp);
            mainMap.MouseDoubleClick += new MouseEventHandler(MainMap_MouseDoubleClick);

            //panel1.Controls.Add(trackBar1);

            mainMap.Overlays.Add(routes);
            mainMap.Overlays.Add(polygons);
            mainMap.Overlays.Add(objects);
            mainMap.Overlays.Add(top);

            routes.Routes.CollectionChanged += new NotifyCollectionChangedEventHandler(Routes_CollectionChanged);
            objects.Markers.CollectionChanged += new NotifyCollectionChangedEventHandler(Markers_CollectionChanged);

            // set current marker
            currentMarker = new GMarkerGoogle(mainMap.Position, GMarkerGoogleType.arrow);
            currentMarker.IsHitTestVisible = false;
            top.Markers.Add(currentMarker);

            //mainMap.VirtualSizeEnabled = true;
            //if(false)
            {
                // add my city location for demo
                GeoCoderStatusCode status = GeoCoderStatusCode.Unknow;
                {
                    PointLatLng? pos = GMapProviders.GoogleMap.GetPoint("Тюмень", out status);
                    if (pos != null && status == GeoCoderStatusCode.G_GEO_SUCCESS)
                    {
                        currentMarker.Position = pos.Value;

                        //GMapMarker myCity = new GMarkerGoogle(pos.Value, GMarkerGoogleType.green_small);
                        //myCity.ToolTipMode = MarkerTooltipMode.Always;
                        //myCity.ToolTipText = "Текст 1 ;}";
                        //objects.Markers.Add(myCity);

                        //139
                        PointLatLng Position1 = new PointLatLng(57.1294124757951, 65.5041468143463);
                        GMapMarker _139 = new GMarkerGoogle(Position1, GMarkerGoogleType.green_small);
                        _139.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _139.ToolTipText = "139";
                        objects.Markers.Add(_139);

                        //75
                        PointLatLng Position2 = new PointLatLng(57.1293367766582, 65.504441857338);
                        GMapMarker _75 = new GMarkerGoogle(Position2, GMarkerGoogleType.green_small);
                        _75.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _75.ToolTipText = "75";
                        objects.Markers.Add(_75);

                        //73
                        PointLatLng Position3 = new PointLatLng(57.1292814579603, 65.5047744512558);
                        GMapMarker _73 = new GMarkerGoogle(Position3, GMarkerGoogleType.gray_small);
                        _73.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _73.ToolTipText = "73 (не подключен)";
                        objects.Markers.Add(_73);

                        //оп1
                        GMapMarker _оп1 = new GMarkerGoogle(new PointLatLng(57.1291679087952, 65.5047208070755), GMarkerGoogleType.blue_small);
                        _оп1.ToolTipMode = MarkerTooltipMode.Always;
                        _оп1.ToolTipText = "оп 1";
                        objects.Markers.Add(_оп1);

                        //71
                        PointLatLng Position4 = new PointLatLng(57.129208670074, 65.5051231384277);
                        GMapMarker _71 = new GMarkerGoogle(Position4, GMarkerGoogleType.orange_small);
                        _71.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _71.ToolTipText = "71 (отключается)";
                        objects.Markers.Add(_71);

                        //69
                        PointLatLng Position5 = new PointLatLng(57.1291300589961, 65.5054396390915);
                        GMapMarker _69 = new GMarkerGoogle(Position5, GMarkerGoogleType.green_small);
                        _69.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _69.ToolTipText = "69";
                        objects.Markers.Add(_69);

                        //67
                        GMapMarker _67 = new GMarkerGoogle(new PointLatLng(57.1290485362204, 65.5057293176651), GMarkerGoogleType.red_small);
                        _67.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _67.ToolTipText = "67 (отключен)";
                        objects.Markers.Add(_67);

                        //65
                        GMapMarker _65 = new GMarkerGoogle(new PointLatLng(57.129004863231, 65.5060940980911), GMarkerGoogleType.blue_small);
                        _65.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _65.ToolTipText = "65 (подключается)";
                        objects.Markers.Add(_65);

                        //63
                        GMapMarker _63 = new GMarkerGoogle(new PointLatLng(57.1289204286389, 65.506437420845), GMarkerGoogleType.yellow_small);
                        _63.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _63.ToolTipText = "63 (другой собственник)";
                        objects.Markers.Add(_63);

                        //61
                        GMapMarker _61 = new GMarkerGoogle(new PointLatLng(57.1288476400428, 65.506796836853), GMarkerGoogleType.green_small);
                        _61.ToolTipMode = MarkerTooltipMode.OnMouseOver;
                        _61.ToolTipText = "61";
                        objects.Markers.Add(_61);


                        GMapOverlay polyOverlay = new GMapOverlay("polygons");

                        List<PointLatLng> points = new List<PointLatLng>();

                        points.Add(new PointLatLng(57.1294124757951, 65.5041468143463));
                        //points.Add(new PointLatLng(57.1293367766582, 65.504441857338));
                        points.Add(new PointLatLng(57.1291679087952, 65.5047208070755));


                        GMapPolygon polygon = new GMapPolygon(points, "mypolygon");
                        //polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                        polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                        polygon.Stroke = new Pen(Color.Red, 1);
                        polyOverlay.Polygons.Add(polygon);

                        mainMap.Overlays.Add(polyOverlay);

                        points = new List<PointLatLng>();

                        points.Add(new PointLatLng(57.1293367766582, 65.504441857338));
                        points.Add(new PointLatLng(57.1291679087952, 65.5047208070755));

                        polygon = new GMapPolygon(points, "mypolygon");
                        //polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                        polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                        polygon.Stroke = new Pen(Color.Red, 1);
                        polyOverlay.Polygons.Add(polygon);

                        points = new List<PointLatLng>();

                        points.Add(new PointLatLng(57.1288476400428, 65.506796836853));
                        points.Add(new PointLatLng(57.1291679087952, 65.5047208070755));

                        polygon = new GMapPolygon(points, "mypolygon");
                        //polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                        polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                        polygon.Stroke = new Pen(Color.Red, 1);
                        polyOverlay.Polygons.Add(polygon);

                        points = new List<PointLatLng>();

                        points.Add(new PointLatLng(57.1291300589961, 65.5054396390915));
                        points.Add(new PointLatLng(57.1291679087952, 65.5047208070755));

                        polygon = new GMapPolygon(points, "mypolygon");
                        //polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                        polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
                        polygon.Stroke = new Pen(Color.Red, 1);
                        polyOverlay.Polygons.Add(polygon);

                        mainMap.Overlays.Add(polyOverlay);

                    }
                }

                // add some points in lithuania
                //AddLocationLithuania("СУЭНКО");
                //AddLocationLithuania("Фидер2");
                //AddLocationLithuania("Фидер3");
                //AddLocationLithuania("Фидер4");

                if (objects.Markers.Count > 0)
                {
                    mainMap.ZoomAndCenterMarkers(null);
                }

                RegeneratePolygon();

            }

            //GMapOverlay markersOverlay = new GMapOverlay("markers");

            //GMarkerGoogle marker = new GMarkerGoogle(new PointLatLng(57.1516535234477, 65.5653762817383),
            //  GMarkerGoogleType.yellow_pushpin);
            //marker.Tag = "Опора";
            //GMapToolTip marekerToolTip = new GMapToolTip(marker);
            //marker.ToolTipMode = MarkerTooltipMode.Always;

            //markersOverlay.Markers.Add(marker);

            //GMarkerGoogle marker1 = new GMarkerGoogle(new PointLatLng(57.1526535234477, 65.5663762817383),
            //  GMarkerGoogleType.yellow_pushpin);
            //marker1.Tag = "Опора2";

            //markersOverlay.Markers.Add(marker1);

            //GMarkerGoogle marker2 = new GMarkerGoogle(new PointLatLng(57.1536535234477, 65.5673762817383),
            //    GMarkerGoogleType.yellow_pushpin);

            //marker2.Tag = "Опора3";

            //markersOverlay.Markers.Add(marker2);

            //mainMap.Overlays.Add(markersOverlay);

            //GMapOverlay polyOverlay = new GMapOverlay("polygons");

            //List<PointLatLng> points = new List<PointLatLng>();

            //points.Add(new PointLatLng(57.1506535234477, 65.5653762817383));
            //points.Add(new PointLatLng(57.1516535234477, 65.5653762817383));
            //points.Add(new PointLatLng(57.1526535234477, 65.5653762817383));
            //points.Add(new PointLatLng(57.1536535234477, 65.5653762817383));

            //points.Add(new PointLatLng(57.1546535234477, 65.5653762817383));
            //points.Add(new PointLatLng(57.1556535234477, 65.5653762817383));
            //points.Add(new PointLatLng(57.1566535234477, 65.5653762817383));
            //points.Add(new PointLatLng(57.1576535234477, 65.5653762817383));

            //GMapPolygon polygon = new GMapPolygon(points, "mypolygon");
            ////polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
            //polygon.Fill = new SolidBrush(Color.FromArgb(50, Color.Red));
            //polygon.Stroke = new Pen(Color.Red, 1);
            //polyOverlay.Polygons.Add(polygon);

            //mainMap.Overlays.Add(polyOverlay);

            //providers.Add(GMapProviders.YandexMap);
            //providers.Add(GMapProviders.YandexSatelliteMap);
            //providers.Add(GMapProviders.YandexHybridMap);
            //providers.Add(GMapProviders.GoogleMap);
            //providers.Add(GMapProviders.OpenStreetMap);
        }

        public GMapMarker AddMarkerMap(double lat, double lng, string markerType = "GMarkerGoogle", int color = 0, int toolTipMode = 2, string toolTipText = "")
        {
            if (markerType == "GMarkerGoogle")
            {
                GMapMarker marker = new GMarkerGoogle(new PointLatLng(lat, lng), (GMarkerGoogleType)color);

                marker.ToolTipMode = (MarkerTooltipMode)toolTipMode;
                marker.ToolTipText = toolTipText;
                return marker;
            }

            return null;
        }

        void RegeneratePolygon()
        {
            List<PointLatLng> polygonPoints = new List<PointLatLng>();

            foreach (GMapMarker m in objects.Markers)
            {
                if (m is GMapMarkerRect)
                {
                    m.Tag = polygonPoints.Count;
                    polygonPoints.Add(m.Position);
                }
            }

            if (polygon == null)
            {
                polygon = new GMapPolygon(polygonPoints, "polygon test");
                polygon.IsHitTestVisible = true;
                polygons.Polygons.Add(polygon);
            }
            else
            {
                polygon.Points.Clear();
                polygon.Points.AddRange(polygonPoints);

                if (polygons.Polygons.Count == 0)
                {
                    polygons.Polygons.Add(polygon);
                }
                else
                {
                    mainMap.UpdatePolygonLocalPosition(polygon);
                }
            }
        }

        void AddLocationLithuania(string place)
        {
            GeoCoderStatusCode status = GeoCoderStatusCode.Unknow;
            PointLatLng? pos = GMapProviders.GoogleMap.GetPoint("Lithuania, " + place, out status);
            if (pos != null && status == GeoCoderStatusCode.G_GEO_SUCCESS)
            {
                GMarkerGoogle m = new GMarkerGoogle(pos.Value, GMarkerGoogleType.green);
                m.ToolTip = new GMapRoundedToolTip(m);

                GMapMarkerRect mBorders = new GMapMarkerRect(pos.Value);
                {
                    mBorders.InnerMarker = m;
                    mBorders.ToolTipText = place;
                    mBorders.ToolTipMode = MarkerTooltipMode.Always;
                }

                objects.Markers.Add(m);
                objects.Markers.Add(mBorders);
            }
        }


        #region -- map events --
        private void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            //mainMap.Zoom = trackBar1.Value / 100.0;
        }

        private void Routes_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            //textBoxrouteCount.Text = routes.Routes.Count.ToString();
        }

        private void Markers_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            //textBoxMarkerCount.Text = objects.Markers.Count.ToString();
        }

        void OnTileCacheComplete()
        {

        }

        void OnTileCacheStart()
        {

        }

        void OnTileCacheProgress(int left)
        {

        }

        void MainMap_OnMarkerLeave(GMapMarker item)
        {
            if (item is GMapMarkerRect)
            {
                CurentRectMarker = null;

                GMapMarkerRect rc = item as GMapMarkerRect;
                rc.Pen.Color = Color.Blue;
            }
        }

        void MainMap_OnMarkerEnter(GMapMarker item)
        {
            ToolTip tt = new ToolTip();

            Random rnd1 = new Random();

            var val = rnd1.Next(new Random().Next());

            tt.Show("Напряжение: " + val.ToString(), (IWin32Window)this.mainMap);



            if (item is GMapMarkerRect)
            {
                GMapMarkerRect rc = item as GMapMarkerRect;
                rc.Pen.Color = Color.Red;

                CurentRectMarker = rc;
            }

        }

        void MainMap_OnPolygonLeave(GMapPolygon item)
        {
            currentPolygon = null;
            item.Stroke.Color = Color.MidnightBlue;
        }

        void MainMap_OnPolygonEnter(GMapPolygon item)
        {
            currentPolygon = item;
            item.Stroke.Color = Color.Red;

        }

        void MainMap_OnRouteLeave(GMapRoute item)
        {
            currentRoute = null;
            item.Stroke.Color = Color.MidnightBlue;
        }

        void MainMap_OnRouteEnter(GMapRoute item)
        {
            currentRoute = item;
            item.Stroke.Color = Color.Red;
        }

        void MainMap_OnMapTypeChanged(GMapProvider type)
        {

        }

        void MainMap_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isMouseDown = false;
            }
        }

        // add demo circle
        void MainMap_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            MapEvent?.Invoke("MainMap_MouseDoubleClick", "");

            //var cc = new GMapMarkerCircle(mainMap.FromLocalToLatLng(e.X, e.Y));
            //objects.Markers.Add(cc);
        }

        void MainMap_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isMouseDown = true;

                if (currentMarker.IsVisible)
                {
                    currentMarker.Position = mainMap.FromLocalToLatLng(e.X, e.Y);

                    //MapEvent?.Invoke("MainMap_MouseDown", currentMarker.Position.Lat.ToString()+";"+currentMarker.Position.Lng.ToString());
                    MapEvent?.Invoke("MainMap_MouseDown", "f58dee4a-8491-11e6-80d9-b499bac04a54");



                    var px = mainMap.MapProvider.Projection.FromLatLngToPixel(currentMarker.Position.Lat, currentMarker.Position.Lng, (int)mainMap.Zoom);
                    var tile = mainMap.MapProvider.Projection.FromPixelToTileXY(px);

                }
            }
        }

        // move current marker with left holding
        void MainMap_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && isMouseDown)
            {
                if (CurentRectMarker == null)
                {
                    if (currentMarker.IsVisible)
                    {
                        currentMarker.Position = mainMap.FromLocalToLatLng(e.X, e.Y);
                    }
                }
                else // move rect marker
                {
                    PointLatLng pnew = mainMap.FromLocalToLatLng(e.X, e.Y);

                    int? pIndex = (int?)CurentRectMarker.Tag;
                    if (pIndex.HasValue)
                    {
                        if (pIndex < polygon.Points.Count)
                        {
                            polygon.Points[pIndex.Value] = pnew;
                            mainMap.UpdatePolygonLocalPosition(polygon);
                        }
                    }

                    if (currentMarker.IsVisible)
                    {
                        currentMarker.Position = pnew;
                    }
                    CurentRectMarker.Position = pnew;

                    if (CurentRectMarker.InnerMarker != null)
                    {
                        CurentRectMarker.InnerMarker.Position = pnew;
                    }
                }

                mainMap.Refresh(); // force instant invalidation
            }
        }

        // MapZoomChanged
        void MainMap_OnMapZoomChanged()
        {
            //trackBar1.Value = (int)(mainMap.Zoom * 100.0);
        }

        // click on some marker
        void MainMap_OnMarkerClick(GMapMarker item, MouseEventArgs e)
        {
            //if (e.Button == System.Windows.Forms.MouseButtons.Left)
            //{
            //if (item is GMapMarkerRect)
            //{
            GeoCoderStatusCode status;
            var pos = GMapProviders.GoogleMap.GetPlacemark(item.Position, out status);
            if (status == GeoCoderStatusCode.G_GEO_SUCCESS && pos != null)
                // {
                //GMapMarkerRect v = item as GMapMarkerRect;
                //{
                item.ToolTipText = pos.Value.Address;
            //      }
            mainMap.Invalidate(false);
            //  }
            //}
            //else
            //  {

            //    }
            // }
        }

        // loader start loading tiles
        void MainMap_OnTileLoadStart()
        {

        }

        // loader end loading tiles
        void MainMap_OnTileLoadComplete(long ElapsedMilliseconds)
        {

        }

        // current point changed
        void MainMap_OnPositionChanged(PointLatLng point)
        {

        }

        #endregion

        #endregion
    }
}