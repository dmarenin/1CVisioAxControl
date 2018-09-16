using GMap.NET.WindowsForms;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace AxControls
{
    [ComVisible(true)]
    //[InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("776E99F1-93E7-42FF-AA1E-982E334A101F")]
    public interface IAxProxyVisio
    {
        #region AxControls

        string Owner1c { get; set; }
        string PathFile { get; set; }
        string Version { get; }

        object AxDrawingControl { get; }

        PictureBox PictureLoading { get; }

        Visio.Page GetActivePage(string PageHeight, string PageWidth);

        Visio.Application GetApplication();

        Visio.Master GetMaster(Visio.Document currentDoc, string type);

        Visio.Cell GetCell(Visio.Shape shape, string propertyName);

        void Message(string mes = "", string cap = "Ошибка");
        void AddCustomProperty(Visio.Shape shape, string PropertyName, string propertyValue);
        void ConnectShapes(object shape1, object shape2, Visio.Shape connector);
        void ActivateShape(Visio.Shape shape);
        void SetEventDblClick(Visio.Shape shape);
        void Clean();
        void Dispose();
        void Undo();
        void Redo();
        void PrintDocument();
        void SendDocument();
        void SetValuesProperties(dynamic structShape);

        string GetValueProperty(Visio.Shape shape, string propertyName);

        bool SetValueProperty(Visio.Shape shape, string propertyName, string propertyValue);

        #endregion

        #region GMap

        string id_YandexHybrid { get; }
        string id_Yandex { get; }
        string id_YandexSatellite { get; }
        string id_GoogleSatellite { get; }
        string id_GoogleHybrid { get; }
        string id_GoogleTerrain { get; }
        string id_Empty { get; }
        string id_OpenStreetMap { get; }
        string id_Google { get; }

        Map MainMap { get; }

        bool MapVisible { get; set; }
        bool MapEnabled { get; set; }

        GMapOverlay MarkerObjects { get; }
        GMapOverlay MarekerRoutes { get; }
        GMapOverlay MarekerPolygons { get; }

        void ChangeMapType(string id);

        void AddMarkers(object overlay, object val);

        GMapMarker AddMarkerMap(double lat, double lng, string markerType = "GMarkerGoogle", int color = 0, int toolTipMode = 2, string toolTipText = "");

        #endregion
    }

    [ComVisible(true)]
    [Guid("76BBC602-9CBD-40b4-A210-CBB844E7AA70")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IAxProxyVisioEvents
    {
        [DispId(1)]
        void MapEvent(string nameEvent, string args);
    }
}