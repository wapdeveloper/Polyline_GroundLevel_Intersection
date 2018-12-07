using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using DotNetARX;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;


namespace Polylines
{
    public class Polylines
    {
        [CommandMethod("AddPolyline")]
        public void AddPolyline()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db=HostApplicationServices.WorkingDatabase;
            using (Transaction trans=db.TransactionManager.StartTransaction())
            {
                string[] alllines=null;

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "文本文件|*.txt|所有文件|*.*";
                string targetpath;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    targetpath = ofd.FileName;
                    alllines = File.ReadAllLines(ofd.FileName);
                }

                int num = alllines.Length - 2;

                Point3dCollection p3c = new Point3dCollection();
                for (int i = 1; i < alllines.Length-1; i++)
                {
                    string[] position = alllines[i].Split(',');
                    Point3d point = new Point3d(double.Parse(position[0]), double.Parse(position[1]), double.Parse(position[2]));
                    p3c.Add(point);
                }

                //创建直线
                Polyline3d pline = new Polyline3d(Poly3dType.SimplePoly, p3c, false);

                //添加对象到模型空间
                db.AddToModelSpace(pline);
                trans.Commit();
            }
        }

        [CommandMethod("OutPolyline")]
        public void OutPolyline()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {

                //Polyline3d pl = null;
                DBObject pl = null;
                //请求在图形区域选择对象
                PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection();
                //如果提示状态OK，表示已选择对象
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;
                    //遍历选择集内的对象
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        //确认返回的是合法的SelectedObject对象
                        if (acSSObj != null)
                        {
                            //以写打开所选对象
                            Entity acEnt = trans.GetObject(acSSObj.ObjectId,
                            OpenMode.ForWrite) as Entity;
                            pl = trans.GetObject(acSSObj.ObjectId,
                            OpenMode.ForRead);
                            if (acEnt != null)
                            {
                                //将对象颜色修改为绿色
                                acEnt.ColorIndex = 1;
                            }
                        }
                    }
                }


                StringBuilder sbs = new StringBuilder();
                sbs.AppendLine("X,Y,Z");
                int num = 0;

                Polyline3d f3d = pl as Polyline3d;
                Polyline f2d = pl as Polyline;
                if (f3d != null)
                {
                    foreach (ObjectId vId in f3d)
                    {
                        //var yy = item as Point3d ;
                        PolylineVertex3d v3d = trans.GetObject(vId, OpenMode.ForRead) as PolylineVertex3d;
                        sbs.AppendLine(v3d.Position.X.ToString("F4") + "," + v3d.Position.Y.ToString("F4") + "," + v3d.Position.Z.ToString("F4"));
                        num++;
                    }
                }
                if (f2d != null)
                {
                    for (int i = 0; i < f2d.NumberOfVertices; i++)
                    {
                        Point3d point=f2d.GetPoint3dAt(i);
                        sbs.AppendLine(point.X.ToString() + "," + point.Y.ToString() + ","+ point.Z.ToString());
                        num++;
                    }
                }



                string[] vps = Regex.Split(sbs.ToString(),"\r\n");

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "文本文件|*.txt|所有文件|*.*";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string targetpath = sfd.FileName;
                    File.WriteAllLines(targetpath, vps);
                }

                //关闭事务
                //var  uuui=uuu.GetSplitCurves(0);
                trans.Commit();
            }
        }



        [CommandMethod("IntersectionPolyline")]
        public void IntersectionPolyline()
        {
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {

                //第一条多段线
                acDoc.Editor.WriteMessage("\n选择多段线: ");
                DBObject pl = null;
                //请求在图形区域选择对象
                PromptSelectionResult acSSPrompt = acDoc.Editor.GetSelection();
                //如果提示状态OK，表示已选择对象
                if (acSSPrompt.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt.Value;
                    //遍历选择集内的对象
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        //确认返回的是合法的SelectedObject对象
                        if (acSSObj != null)
                        {
                            //以写打开所选对象
                            Entity acEnt = trans.GetObject(acSSObj.ObjectId,
                            OpenMode.ForWrite) as Entity;
                            pl = trans.GetObject(acSSObj.ObjectId,
                            OpenMode.ForRead);
                            if (acEnt != null)
                            {
                                //将对象颜色修改为绿色
                                acEnt.ColorIndex = 1;
                            }
                        }
                    }
                }




                //选择的等高线
                acDoc.Editor.WriteMessage("\n选择相交的等高线: ");
                //等高线集合
                List<DBObject> obs = new List<DBObject>();
                //请求在图形区域选择对象
                PromptSelectionResult  acSSPrompt1 = acDoc.Editor.GetSelection();
                //如果提示状态OK，表示已选择对象
                if (acSSPrompt1.Status == PromptStatus.OK)
                {
                    SelectionSet acSSet = acSSPrompt1.Value;
                    //遍历选择集内的对象
                    foreach (SelectedObject acSSObj in acSSet)
                    {
                        //确认返回的是合法的SelectedObject对象
                        if (acSSObj != null)
                        {
                            //以写打开所选对象
                            Entity acEnt = trans.GetObject(acSSObj.ObjectId,
                            OpenMode.ForWrite) as Entity;
                            DBObject pl2 = trans.GetObject(acSSObj.ObjectId,
                            OpenMode.ForRead);
                            obs.Add(pl2);

                            if (acEnt != null)
                            {
                                //将对象颜色修改为绿色
                                acEnt.ColorIndex = 5;
                            }
                        }
                    }
                }




                //Point3dCollection p3ds = new Point3dCollection(); 
                List<Point3d> p3ds = new List<Point3d>();

                StringBuilder sbs = new StringBuilder();
                sbs.AppendLine("X,Y,Z");
                int num = 0;

                //第一条多段线
                Polyline f2d = pl as Polyline;

                foreach (Polyline highpl in obs)
                {
                    f2d.Elevation = highpl.Elevation;
                    //多段线和等高线的交点点集
                    Point3dCollection pds = new Point3dCollection();
                    f2d.IntersectWith(highpl, Intersect.OnBothOperands, pds, 0, 0);
                    for (int i = 0; i < pds.Count; i++)
                    {
                        p3ds.Add(pds[i]);
                    }
                }


                //按坐标排序
                List<Point3d> points =p3ds.OrderBy(p=>p.X).ToList();
                //判断点的高程，按从高到低生成剖面
                double Z1 = points[0].Z;
                double Z2 = points[points.Count - 1].Z;
                if(Z1<Z2)
                    points = p3ds.OrderByDescending(p => p.X).ToList();

                ///生成剖面——————————————————
                //算基准点
                double x0 = points[0].X;
                double y0 = points[0].Y- points[0].Z;

                //剖面线点集
                Point2dCollection p2c = new Point2dCollection();

                //剖面线上第一个点
                Point2d p2d = new Point2d(points[0].X, points[0].Y);
                p2c.Add(p2d);


                //累计值
                List<double> dx = new List<double>();
                List<double> dH = new List<double>();                
                for (int i = 0; i < points.Count-1; i++)
                {
                    dx.Add( Math.Sqrt((points[i + 1].X - points[i].X) * (points[i + 1].X - points[i].X) + (points[i + 1].Y - points[i].Y) * (points[i + 1].Y - points[i].Y)));
                    dH.Add(points[i + 1].Z);
                }

                //剖面线上的点
                double sdx = x0; //dx累计值
                for (int i = 0; i < dx.Count; i++)
                {
                    sdx += dx[i];
                    p2c.Add(new Point2d(sdx, y0+dH[i]));
                }

                ///画剖面线——————————————————
                Polyline ply2d = new Polyline();
                ply2d.CreatePolyline(p2c);


                //输出交点坐标
                if (f2d != null)
                {
                    for (int i = 0; i < points.Count; i++)
                    {
                        Point3d point = points[i];
                        sbs.AppendLine(point.X.ToString("F4") + "," + point.Y.ToString("F4") + "," + point.Z.ToString("F4"));
                        num++;
                    }
                }

                string[] vps = Regex.Split(sbs.ToString(), "\r\n");
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "文本文件|*.txt|所有文件|*.*";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    string targetpath = sfd.FileName;
                    File.WriteAllLines(targetpath, vps);
                }

                //添加对象到模型空间
                db.AddToModelSpace(ply2d);
                trans.Commit();
            }
        }
    }
}
