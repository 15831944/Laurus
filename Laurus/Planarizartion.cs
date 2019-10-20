using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
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
        private static Point3d ClearZ(this Point3d inputPoint) {
            Point3d tmpPoint = new Point3d(inputPoint.X, inputPoint.Y, 0.0);
            return tmpPoint;
        }

        /// <summary>
        ///     清除参数的Z方向矢量并返回
        /// </summary>
        /// <param name="inputVector">待修改的三维矢量</param>
        /// <returns>修改后的三维矢量</returns>
        private static Vector3d ClearZ(this Vector3d inputVector) {
            Vector3d tmpVector = new Vector3d(inputVector.X, inputVector.Y, 0.0);
            return tmpVector;
        }

        /*
        *---------------------------------------------------------------------------
        *    以下是各类元素的清零函数
        */
        #region 各类元素的清零函数
        public static bool Clear(this Line obj) {
            if (Math.Abs(obj.StartPoint.Z) < TOLERANCE && Math.Abs(obj.EndPoint.Z) < TOLERANCE) {
                return false;
            }
            obj.StartPoint = obj.StartPoint.ClearZ();
            obj.EndPoint = obj.EndPoint.ClearZ();
            return true;
        }

        public static bool Clear(this Xline obj) {
            if (Math.Abs(obj.BasePoint.Z) < TOLERANCE && Math.Abs(obj.SecondPoint.Z) < TOLERANCE) {
                return false;
            }
            obj.BasePoint = obj.BasePoint.ClearZ();
            obj.SecondPoint = obj.SecondPoint.ClearZ();
            return true;
        }

        public static bool Clear(this Ray obj) {
            if (Math.Abs(obj.BasePoint.Z) < TOLERANCE && Math.Abs(obj.SecondPoint.Z) < TOLERANCE) {
                return false;
            }
            obj.BasePoint = obj.BasePoint.ClearZ();
            obj.SecondPoint = obj.SecondPoint.ClearZ();
            return true;
        }

        public static bool Clear(this Arc obj) {
            if (Math.Abs(obj.Center.Z) < TOLERANCE) {
                return false;
            }
            obj.Center = obj.Center.ClearZ();
            return true;
        }

        public static bool Clear(this Circle obj) {
            if (Math.Abs(obj.Center.Z) < TOLERANCE) {
                return false;
            }
            obj.Center = obj.Center.ClearZ();
            return true;
        }

        public static bool Clear(this Ellipse obj) {
            if (Math.Abs(obj.Center.Z) < TOLERANCE) {
                return false;
            }
            obj.Center = obj.Center.ClearZ();
            return true;
        }

        public static bool Clear(this DBText obj) {
            if (Math.Abs(obj.Position.Z) < TOLERANCE) {
                return false;
            }
            obj.Position = obj.Position.ClearZ();
            return true;
        }

        public static bool Clear(this MText obj) {
            if (Math.Abs(obj.Location.Z) < TOLERANCE && Math.Abs(obj.Direction.Z) < TOLERANCE) {
                return false;
            }
            obj.Location = obj.Location.ClearZ();
            obj.Direction = obj.Direction.ClearZ();
            return true;
        }

        public static bool Clear(this Polyline obj) {
            if (Math.Abs(obj.Elevation) < TOLERANCE) {
                return false;
            }
            obj.Elevation = 0.0;
            return true;
        }

        public static bool Clear(this Polyline2d obj) {
            if (Math.Abs(obj.Elevation) < TOLERANCE) {
                return false;
            }
            obj.Elevation = 0.0;
            return true;
        }

        public static bool Clear(this DBPoint obj) {
            if (Math.Abs(obj.Position.Z) < TOLERANCE) {
                return false;
            }
            obj.Position = obj.Position.ClearZ();
            return true;
        }

        public static bool Clear(this Dimension obj) {
            bool myFlag = false;
            if (Math.Abs(obj.Elevation) >= TOLERANCE) {
                obj.Elevation = 0.0;
                myFlag = true;
            }
            if (obj is AlignedDimension) {
                AlignedDimension myDimension = obj as AlignedDimension;
                if (Math.Abs(myDimension.XLine1Point.Z) >= TOLERANCE) {
                    myDimension.XLine1Point = myDimension.XLine1Point.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine2Point.Z) >= TOLERANCE) {
                    myDimension.XLine2Point = myDimension.XLine2Point.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.DimLinePoint.Z) >= TOLERANCE) {
                    myDimension.DimLinePoint = myDimension.DimLinePoint.ClearZ();
                    myFlag = true;
                }
            } else if (obj is ArcDimension) {
                ArcDimension myDimension = obj as ArcDimension;
                if (Math.Abs(myDimension.XLine1Point.Z) >= TOLERANCE) {
                    myDimension.XLine1Point = myDimension.XLine1Point.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine2Point.Z) >= TOLERANCE) {
                    myDimension.XLine2Point = myDimension.XLine2Point.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.Leader1Point.Z) >= TOLERANCE) {
                    myDimension.Leader1Point = myDimension.Leader1Point.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.Leader2Point.Z) >= TOLERANCE) {
                    myDimension.Leader2Point = myDimension.Leader2Point.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.ArcPoint.Z) >= TOLERANCE) {
                    myDimension.ArcPoint = myDimension.ArcPoint.ClearZ();
                    myFlag = true;
                }
            } else if (obj is DiametricDimension) {
                DiametricDimension myDimension = obj as DiametricDimension;
                if (Math.Abs(myDimension.ChordPoint.Z) >= TOLERANCE) {
                    myDimension.ChordPoint = myDimension.ChordPoint.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.FarChordPoint.Z) >= TOLERANCE) {
                    myDimension.FarChordPoint = myDimension.FarChordPoint.ClearZ();
                    myFlag = true;
                }
            } else if (obj is LineAngularDimension2) {
                LineAngularDimension2 myDimension = obj as LineAngularDimension2;
                if (Math.Abs(myDimension.ArcPoint.Z) >= TOLERANCE) {
                    myDimension.ArcPoint = myDimension.ArcPoint.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine1End.Z) >= TOLERANCE) {
                    myDimension.XLine1End = myDimension.XLine1End.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine1Start.Z) >= TOLERANCE) {
                    myDimension.XLine1Start = myDimension.XLine1Start.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine2End.Z) >= TOLERANCE) {
                    myDimension.XLine2End = myDimension.XLine2End.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine2Start.Z) >= TOLERANCE) {
                    myDimension.XLine2Start = myDimension.XLine2Start.ClearZ();
                    myFlag = true;
                }
            } else if (obj is Point3AngularDimension) {
                Point3AngularDimension myDimension = obj as Point3AngularDimension;
                if (Math.Abs(myDimension.ArcPoint.Z) >= TOLERANCE) {
                    myDimension.ArcPoint = myDimension.ArcPoint.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.CenterPoint.Z) >= TOLERANCE) {
                    myDimension.CenterPoint = myDimension.CenterPoint.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine1Point.Z) >= TOLERANCE) {
                    myDimension.XLine1Point = myDimension.XLine1Point.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine2Point.Z) >= TOLERANCE) {
                    myDimension.XLine2Point = myDimension.XLine2Point.ClearZ();
                    myFlag = true;
                }
            } else if (obj is RadialDimension) {
                RadialDimension myDimension = obj as RadialDimension;
                if (Math.Abs(myDimension.Center.Z) >= TOLERANCE) {
                    myDimension.Center = myDimension.Center.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.ChordPoint.Z) >= TOLERANCE) {
                    myDimension.ChordPoint = myDimension.ChordPoint.ClearZ();
                    myFlag = true;
                }
            } else if (obj is RadialDimensionLarge) {
                RadialDimensionLarge myDimension = obj as RadialDimensionLarge;
                if (Math.Abs(myDimension.Center.Z) >= TOLERANCE) {
                    myDimension.Center = myDimension.Center.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.ChordPoint.Z) >= TOLERANCE) {
                    myDimension.ChordPoint = myDimension.ChordPoint.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.JogPoint.Z) >= TOLERANCE) {
                    myDimension.JogPoint = myDimension.JogPoint.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.OverrideCenter.Z) >= TOLERANCE) {
                    myDimension.OverrideCenter = myDimension.OverrideCenter.ClearZ();
                    myFlag = true;
                }
            } else if (obj is RotatedDimension) {
                RotatedDimension myDimension = obj as RotatedDimension;
                if (Math.Abs(myDimension.DimLinePoint.Z) >= TOLERANCE) {
                    myDimension.DimLinePoint = myDimension.DimLinePoint.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine1Point.Z) >= TOLERANCE) {
                    myDimension.XLine1Point = myDimension.XLine1Point.ClearZ();
                    myFlag = true;
                }
                if (Math.Abs(myDimension.XLine2Point.Z) >= TOLERANCE) {
                    myDimension.XLine2Point = myDimension.XLine2Point.ClearZ();
                    myFlag = true;
                }
            }
            return myFlag;
        }

        public static bool Clear(this Hatch obj) {
            if (Math.Abs(obj.Elevation) < TOLERANCE) {
                return false;
            }
            obj.Elevation = 0.0;
            return true;
        }

        public static bool Clear(this Spline obj) {
            //将Splin的ControlPoints赋值到数组，并检查Z坐标
            bool myFlag = false;
            Point3d[] objControlPoints = new Point3d[obj.NumControlPoints];
            for (int i = 0; i < objControlPoints.Length; i++) {
                objControlPoints[i] = obj.GetControlPointAt(i);
                if (Math.Abs(objControlPoints[i].Z) > TOLERANCE) {
                    obj.SetControlPointAt(i, objControlPoints[i]);
                }
            }
            myFlag = true;
            return myFlag;
        }

        public static bool Clear(this Solid obj) {
            bool myFlag = false;
            Point3d[] objControlPoints = new Point3d[4];
            for (short i = 0; i < 4; i++) {
                objControlPoints[i] = obj.GetPointAt(i);
                if (Math.Abs(objControlPoints[i].Z) > TOLERANCE) {
                    obj.SetPointAt(i, objControlPoints[i]);
                }
            }
            myFlag = true;
            return myFlag;
        }

        public static bool Clear(this BlockReference obj) {
            if (Math.Abs(obj.Position.Z) < TOLERANCE) {
                return false;
            }
            obj.Position = obj.Position.ClearZ();
            return true;
        }
        #endregion
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
            Database acCurDb = HostApplicationServices.WorkingDatabase;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor acEdt = acDoc.Editor;

            if (cEntity is Line) {
                return ((Line)cEntity).Clear();
            } else if (cEntity is Xline) {
                return ((Xline)cEntity).Clear();
            } else if (cEntity is Ray) {
                return ((Ray)cEntity).Clear();
            } else if (cEntity is Arc) {
                return ((Arc)cEntity).Clear();
            } else if (cEntity is Circle) {
                return ((Circle)cEntity).Clear();
            } else if (cEntity is Ellipse) {
                return ((Ellipse)cEntity).Clear();
            } else if (cEntity is DBText) {
                return ((DBText)cEntity).Clear();
            } else if (cEntity is MText) {
                return ((MText)cEntity).Clear();
            } else if (cEntity is Polyline) {
                return ((Polyline)cEntity).Clear();
            } else if (cEntity is Polyline2d) {
                return ((Polyline2d)cEntity).Clear();
            } else if (cEntity is DBPoint) {
                return ((DBPoint)cEntity).Clear();
            } else if (cEntity is Dimension) {
                return ((Dimension)cEntity).Clear();
            } else if (cEntity is Hatch) {
                return ((Hatch)cEntity).Clear();
            } else if (cEntity is Spline) {
                return ((Spline)cEntity).Clear();
            } else if (cEntity is Solid) {
                return ((Solid)cEntity).Clear();
            } else if (cEntity is BlockReference) {
                return ((BlockReference)cEntity).Clear();
            } else {
                acEdt.WriteMessage("发现了一个未知图元类型：" + cEntity.GetType().ToString() + "。\n");
                return false;
            }
        }

        /// <summary>
        ///     对传入的BlockTableRecord中的每个图元单独运行清零函数
        /// </summary>
        /// <param name="btRecord">待清零的块定义</param>
        /// <param name="countTotal">检查过的图元计数</param> 
        /// <param name="countCleared">清零的图元计数</param> 
        /// <param name="recurseIn">是否清零块内和锁定图层内的图元的Z坐标，true：清除，false：不清除</param> 
        /// <returns></returns>
        internal static void Clear(this BlockTableRecord btRecord, ref ulong countTotal, ref ulong countCleared, bool recurseIn = false) {
            Database acCurDb = HostApplicationServices.WorkingDatabase;
            List<Entity> objList = new List<Entity>();
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction()) {
                foreach (ObjectId id in btRecord) {
                    Entity myObj = (Entity)acTrans.GetObject(id, OpenMode.ForWrite, false, true);
                    LayerTable acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                    LayerTableRecord acLyrTblRec = acTrans.GetObject(myObj.LayerId, OpenMode.ForWrite) as LayerTableRecord;
                    bool curLayerLocked = acLyrTblRec.IsLocked;
                    acLyrTblRec.IsLocked = false;
                    //对块内和锁定图层内的图元进行清零操作，递归
                    if (myObj is BlockReference && recurseIn) {
                        BlockReference br = acTrans.GetObject(id, OpenMode.ForRead) as BlockReference;
                        BlockTableRecord myBlock = (BlockTableRecord)br.BlockTableRecord.GetObject(OpenMode.ForWrite, false, true);
                        myBlock.Clear(ref countTotal, ref countCleared, recurseIn);
                    }
                    //对普通图元进行操作
                    if (!curLayerLocked || recurseIn) {
                        if (myObj.Clear()) {
                            countCleared++;
                        }
                    }
                    countTotal++;
                    acLyrTblRec.IsLocked = curLayerLocked;
                }
                acTrans.Commit();
            }
            return;
        }

        /// <summary>
        ///     接收命令和参数，遍历模型空间清零并计数。
        /// </summary>
        [CommandMethod("clear")]
        public static void Clear() {
            Database acCurDb = HostApplicationServices.WorkingDatabase;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor acEdt = acDoc.Editor;

            /// 让用户选择是否清除块内的图元，Y：清除，N：不清除
            PromptStringOptions optRecurseIn = new PromptStringOptions("是否包括块内和锁定图层内的图元？[是(Y)/否(N)]") {
                AllowSpaces = false,
                DefaultValue = "N"
            };
            PromptResult resStringResult;
            do {
                resStringResult = acEdt.GetString(optRecurseIn);
            } while (resStringResult.StringResult.ToUpper() != "Y" && resStringResult.StringResult.ToUpper() != "N");
            bool recurseIn = (resStringResult.StringResult.ToUpper() == "Y");

            //对图形进行修改
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction()) {
                ulong countTotal = 0;
                ulong countCleared = 0;
                //遍历模型空间的清零
                BlockTableRecord btRecord = (BlockTableRecord)acTrans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(acCurDb), OpenMode.ForRead, false, true);
                btRecord.Clear(ref countTotal, ref countCleared, recurseIn);
                acEdt.WriteMessage("共检查了" + countTotal + "个对象，其中" + countCleared + "个对象的Z坐标已经清除。\n");
                //提交执行
                acTrans.Commit();
            }
        }

        /// <summary>
        ///     用于测试
        /// </summary>
        [CommandMethod("test")]
        public static void Test() {
            Database acCurDb = HostApplicationServices.WorkingDatabase;
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Editor acEdt = acDoc.Editor;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction()) {

            }
        }
    }

}