using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using DotNetARX;
using Exception = Autodesk.AutoCAD.Runtime.Exception;

namespace Laurus {

    /// <summary>
    ///     提供clear命令
    ///     遍历文档中的所有对象，将所有Z坐标不等于0的元素清零
    /// </summary>
    internal static class Planarizartion {
        private const double TOLERANCE = 0.0001;

        /// <summary>
        ///     清除参数的Z坐标并返回
        /// </summary>
        /// <param name="inputPoint">待修改的三维坐标</param>
        /// <returns>修改后的三维坐标</returns>
        private static Point3d ClearZ(Point3d inputPoint) {
            Point3d tmpPoint = new Point3d(inputPoint.X, inputPoint.Y, 0.0);
            return tmpPoint;
        }

        /// <summary>
        ///     清除参数的Z方向矢量并返回
        /// </summary>
        /// <param name="inputVector">待修改的三维矢量</param>
        /// <returns>修改后的三维矢量</returns>
        private static Vector3d ClearZ(Vector3d inputVector) {
            Vector3d tmpVector = new Vector3d(inputVector.X, inputVector.Y, 0.0);
            return tmpVector;
        }

        /*
        *---------------------------------------------------------------------------
        *    以下是各类元素的清零函数
        */

        public static bool Clear(this Line obj) {
            if (Math.Abs(obj.StartPoint.Z) < TOLERANCE && Math.Abs(obj.EndPoint.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.StartPoint = ClearZ(obj.StartPoint);
            obj.EndPoint = ClearZ(obj.EndPoint);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Xline obj) {
            if (Math.Abs(obj.BasePoint.Z) < TOLERANCE && Math.Abs(obj.SecondPoint.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.BasePoint = ClearZ(obj.BasePoint);
            obj.SecondPoint = ClearZ(obj.SecondPoint);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Ray obj) {
            if (Math.Abs(obj.BasePoint.Z) < TOLERANCE && Math.Abs(obj.SecondPoint.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.BasePoint = ClearZ(obj.BasePoint);
            obj.SecondPoint = ClearZ(obj.SecondPoint);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Arc obj) {
            if (Math.Abs(obj.Center.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Center = ClearZ(obj.Center);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Circle obj) {
            if (Math.Abs(obj.Center.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Center = ClearZ(obj.Center);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Ellipse obj) {
            if (Math.Abs(obj.Center.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Center = ClearZ(obj.Center);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this DBText obj) {
            if (Math.Abs(obj.Position.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Position = ClearZ(obj.Position);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this MText obj) {
            if (Math.Abs(obj.Location.Z) < TOLERANCE && Math.Abs(obj.Direction.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Location = ClearZ(obj.Location);
            obj.Direction = ClearZ(obj.Direction);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Polyline obj) {
            if (Math.Abs(obj.Elevation) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Elevation = 0.0;
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this DBPoint obj) {
            if (Math.Abs(obj.Position.Z) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Position = ClearZ(obj.Position);
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Dimension obj) {
            if (Math.Abs(obj.Elevation) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Elevation = 0.0;
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Hatch obj) {
            if (Math.Abs(obj.Elevation) < TOLERANCE) return false;
            obj.UpgradeOpen();
            obj.Elevation = 0.0;
            obj.DowngradeOpen();
            return true;
        }

        public static bool Clear(this Spline obj) {
            //将Splin的ControlPoints赋值到数组，并检查Z坐标
            bool myFlag = false;
            Point3d[] objControlPoints = new Point3d[obj.NumControlPoints];
            for (int i = 0; i < objControlPoints.Length; i++) {
                objControlPoints[i] = obj.GetControlPointAt(i);
                if (Math.Abs(objControlPoints[i].Z) > TOLERANCE) myFlag = true;
            }

            //根据检查结果进行处理
            if (!myFlag) return false;
            obj.UpgradeOpen();
            for (int i = 0; i < objControlPoints.Length; i++) {
                obj.SetControlPointAt(i, ClearZ(objControlPoints[i]));
            }
            obj.DowngradeOpen();
            return true;
        }

        /*
        *    以上是各类元素的清零函数
        *---------------------------------------------------------------------------
        */

        /// <summary>
        ///     单个实体的清零函数，按类型分发调用各类型的清零函数，返回值用于主函数计数。
        /// </summary>
        /// <param name="cEntity">待清零的实体</param>
        /// <returns>返回该实体是否需清零，以便计数。true：需要清零，false：不需要清零。</returns>
        public static bool Clear(this Entity cEntity) {
            if (cEntity is Line) return ((Line) cEntity).Clear();
            if (cEntity is Xline) return ((Xline) cEntity).Clear();
            if (cEntity is Ray) return ((Ray) cEntity).Clear();
            if (cEntity is Arc) return ((Arc) cEntity).Clear();
            if (cEntity is Circle) return ((Circle) cEntity).Clear();
            if (cEntity is Ellipse) return ((Ellipse) cEntity).Clear();
            if (cEntity is DBText) return ((DBText) cEntity).Clear();
            if (cEntity is MText) return ((MText) cEntity).Clear();
            if (cEntity is Polyline) return ((Polyline) cEntity).Clear();
            if (cEntity is DBPoint) return ((DBPoint) cEntity).Clear();
            if (cEntity is Dimension) return ((Dimension) cEntity).Clear();
            if (cEntity is Hatch) return ((Hatch) cEntity).Clear();
            if (cEntity is Spline) return ((Spline) cEntity).Clear();
            return false;
        }

        /// <summary>
        ///     对传入List中的每个元素单独运行清零函数，如遇到块，则根据参数2选择是否递归清零
        /// </summary>
        /// <param name="objList">待清零的图元列表</param>
        /// <param name="recurseIn">是否递归清除块内的图元，true：清除，false：不清除块内图元</param>
        /// <returns></returns>
        internal static ulong Clear(List<Entity> objList, bool recurseIn) {
            ulong countCleared = 0; //清零计数
            foreach (Entity myObj in objList) {
                if (recurseIn && myObj is BlockReference) {
                    //块内图元处理过程
                }
                if (myObj.Clear()) countCleared++;
            }
            return countCleared;
        }

        /// <summary>
        ///     接收命令和参数，遍历模型空间清零并计数。
        /// </summary>
        [CommandMethod("clear")]
        public static void Clear() {
            Database db = HostApplicationServices.WorkingDatabase;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptStringOptions optRecurseIn = new PromptStringOptions("是否包括块内的图元？[是(Y)/否(N)]") {
                AllowSpaces = false,
                DefaultValue = "N"
            };
            PromptResult resStringResult;
            do {
                resStringResult = ed.GetString(optRecurseIn);
            } while (resStringResult.StringResult.ToUpper() != "Y" && resStringResult.StringResult.ToUpper() != "N");
            bool recurseIn = (resStringResult.StringResult == "Y");
            using (Transaction trans = db.TransactionManager.StartTransaction()) {
                try {
                    List<Entity> objList = db.GetEntsInDatabase();
                    ulong countTotal = (ulong) objList.Count;
                    ulong countCleared = Clear(objList, recurseIn);

                    ed.WriteMessageWithReturn(
                        "共检查了" + countTotal + "个对象，其中" + countCleared + "个对象的Z坐标已经清除。");
                    trans.Commit();
                } catch (Exception e) {
                    trans.Abort();
                    ed.WriteMessageWithReturn(e.ToString());
                }
            }
        }

        /// <summary>
        ///     用于测试
        /// </summary>
        [CommandMethod("test")]
        public static void Test() {
            Database db = HostApplicationServices.WorkingDatabase;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            List<Entity> objList = db.GetEntsInDatabase();
            ed.WriteMessageWithReturn("模型空间共有" + objList.Count + "个对象。");
        }
    }

}