using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Laurus;

[assembly: ExtensionApplication(typeof (Init))]
[assembly: CommandClass(typeof (Init))]
//[assembly: CommandClass(typeof (Planarizartion))]

namespace Laurus {
    internal class Init : IExtensionApplication {
        public void Initialize() {
            Hello();
        }

        public void Terminate() {
        }

        [CommandMethod("hello")]
        public void Hello() {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("欢迎使用Laurus系统！");
        }
    }
}