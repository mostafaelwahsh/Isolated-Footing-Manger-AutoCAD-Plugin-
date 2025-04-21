using Autodesk.AutoCAD.Runtime;


using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using System.Windows;

[assembly: CommandClass(typeof(CADAPI.Commands.FootingWindow))]
namespace CADAPI.Commands
{
    public class FootingWindow
    {
        [CommandMethod("IFM")]
        public void ShowFootingUI()
        {
            var window = new FootingManger();
            var helper = new System.Windows.Interop.WindowInteropHelper(window);
            helper.Owner = Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.Handle;
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalWindow(window);
        }
    }
}
