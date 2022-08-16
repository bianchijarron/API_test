using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Diagnostics;

namespace API_test
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class1:IExternalCommand
    {
        static string stationData = File.ReadAllText(@"E:\StationData.txt");
        static JObject JB_stationData = (JObject)JsonConvert.DeserializeObject(stationData);
        static string levelData = File.ReadAllText(@"E:\LevelData.txt");
        static JArray JB_levelData = (JArray)JsonConvert.DeserializeObject(levelData);
        static string digData = File.ReadAllText(@"E:\DigsData.txt");
        static JArray JB_digData = (JArray)JsonConvert.DeserializeObject(digData);

        public double station_centerX = Convert.ToDouble(JB_stationData["centerX"].ToString()) / 0.3048;
        public double station_centerY = Convert.ToDouble(JB_stationData["centerY"].ToString()) / 0.3048;
        public double station_length = Convert.ToDouble(JB_stationData["length"].ToString()) / 0.3048;
        public double station_width = Convert.ToDouble(JB_stationData["width"].ToString()) / 0.3048;
        public double station_left = Convert.ToDouble(JB_stationData["left"].ToString()) / 0.3048;
        public double station_right = Convert.ToDouble(JB_stationData["right"].ToString()) / 0.3048;
        public double station_top = Convert.ToDouble(JB_stationData["top"].ToString()) / 0.3048;
        public double station_bottom = Convert.ToDouble(JB_stationData["bottom"].ToString()) / 0.3048;
        public double station_c_wall = Convert.ToDouble(JB_stationData["c_wall"].ToString()) / 0.3048;
        public double station_s_wall = Convert.ToDouble(JB_stationData["s_wall"].ToString()) / 0.3048;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet element)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementCategoryFilter myfilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralColumns);
            ICollection<Element> levels = collector.OfClass(typeof(Level)).ToElements();
            FilteredElementCollector walls = new FilteredElementCollector(doc).OfClass(typeof(Wall));
            FilteredElementCollector floors = new FilteredElementCollector(doc).OfClass(typeof(Floor));
            List<ElementId> wallList = new List<ElementId>();
            List<ElementId> levelList = new List<ElementId>();
            List<ElementId> floorList = new List<ElementId>();
            List<Line> s_wallEdgeLine = new List<Line>();
            List<Line> c_wallEdgeLine = new List<Line>();

            using (Transaction ttt = new Transaction(doc, "Delete Model"))
            {
                ttt.Start();

                wallList = new List<ElementId>();//----------------------刪除牆
                if (walls.Count<Element>() != 0)
                {
                    MessageBox.Show(walls.Count<Element>().ToString());

                    foreach (Wall elm in walls)
                    {
                        wallList.Add(elm.Id);
                    }
                    doc.Delete(wallList);
                }

                if (floors.Count<Element>() != 0)//----------------------刪除樓板
                {
                    foreach (Floor elm in floors)
                    {
                        floorList.Add(elm.Id);
                    }
                    doc.Delete(floorList);
                }

                if (levels.Count != 0)//----------------------刪除樓層
                {
                    foreach (Element elm in levels)
                    {
                        levelList.Add(elm.Id);
                    }
                    doc.Delete(levelList);
                }

                ttt.Commit();
            }

            for (int i = 0; i < JB_levelData.Count; i++)//----------------------新增樓層
            {
                if (JB_levelData[i]["selected"].ToString() == "True")
                {
                    double level_height = Convert.ToDouble(JB_levelData[i]["height"].ToString());

                    using (Transaction ttt = new Transaction(doc, "Create Level API"))
                    {
                        ttt.Start();

                        Level level = Level.Create(doc, level_height / 0.3048);
                        level.Name = JB_levelData[i]["name"].ToString();
                        //levels.Add(level);

                        ttt.Commit();
                    }
                }
                
            }
            
            string NewWallTypeName_s_wall = "站體結構牆";
            string NewWallTypeName_c_wall = "站體連續壁";
            WallType NewWallType_s_wall = null;
            WallType NewWallType_c_wall = null;

            collector = new FilteredElementCollector(doc).OfClass(typeof(WallType));

            XYZ point_ST = new XYZ(0, 0, 0);
            XYZ point_stationCorner01 = new XYZ(station_centerX  + - station_length / 2 - station_left, station_centerY + station_width / 2 + station_top, 0);
            XYZ point_stationCorner02 = new XYZ(station_centerX  + - station_length / 2 - station_left, station_centerY + - station_width / 2 - station_bottom, 0);
            XYZ point_stationCorner03 = new XYZ(station_centerX + station_length / 2 + station_right, station_centerY + station_width / 2 + station_top, 0);
            XYZ point_stationCorner04 = new XYZ(station_centerX + station_length / 2 + station_right, station_centerY + - station_width / 2 - station_bottom, 0);

            Line Edge01_s_wall = Line.CreateBound(new XYZ(point_stationCorner01.X - station_s_wall / 2, point_stationCorner01.Y + station_s_wall / 2, point_stationCorner01.Z), new XYZ(point_stationCorner02.X - station_s_wall / 2, point_stationCorner02.Y - station_s_wall / 2, point_stationCorner02.Z));
            Line Edge02_s_wall = Line.CreateBound(new XYZ(point_stationCorner01.X - station_s_wall / 2, point_stationCorner01.Y + station_s_wall / 2, point_stationCorner01.Z), new XYZ(point_stationCorner03.X + station_s_wall / 2, point_stationCorner03.Y + station_s_wall / 2, point_stationCorner03.Z));
            Line Edge03_s_wall = Line.CreateBound(new XYZ(point_stationCorner03.X + station_s_wall / 2, point_stationCorner03.Y + station_s_wall / 2, point_stationCorner03.Z), new XYZ(point_stationCorner04.X + station_s_wall / 2, point_stationCorner04.Y - station_s_wall / 2, point_stationCorner04.Z));
            Line Edge04_s_wall = Line.CreateBound(new XYZ(point_stationCorner02.X - station_s_wall / 2, point_stationCorner02.Y - station_s_wall / 2, point_stationCorner02.Z), new XYZ(point_stationCorner04.X + station_s_wall / 2, point_stationCorner04.Y - station_s_wall / 2, point_stationCorner04.Z));
            s_wallEdgeLine = new List<Line>();
            s_wallEdgeLine.Add(Edge01_s_wall);
            s_wallEdgeLine.Add(Edge02_s_wall);
            s_wallEdgeLine.Add(Edge03_s_wall);
            s_wallEdgeLine.Add(Edge04_s_wall);

            Line Edge01_c_wall = Line.CreateBound(new XYZ(point_stationCorner01.X - station_s_wall - station_c_wall / 2, point_stationCorner01.Y + station_s_wall + station_c_wall / 2, point_stationCorner01.Z), new XYZ(point_stationCorner02.X - station_s_wall - station_c_wall / 2, point_stationCorner02.Y - station_s_wall - station_c_wall / 2, point_stationCorner02.Z));
            Line Edge02_c_wall = Line.CreateBound(new XYZ(point_stationCorner01.X - station_s_wall - station_c_wall / 2, point_stationCorner01.Y + station_s_wall + station_c_wall / 2, point_stationCorner01.Z), new XYZ(point_stationCorner03.X + station_s_wall + station_c_wall / 2, point_stationCorner03.Y + station_s_wall + station_c_wall / 2, point_stationCorner03.Z));
            Line Edge03_c_wall = Line.CreateBound(new XYZ(point_stationCorner03.X + station_s_wall + station_c_wall / 2, point_stationCorner03.Y + station_s_wall + station_c_wall / 2, point_stationCorner03.Z), new XYZ(point_stationCorner04.X + station_s_wall + station_c_wall / 2, point_stationCorner04.Y - station_s_wall - station_c_wall / 2, point_stationCorner04.Z));
            Line Edge04_c_wall = Line.CreateBound(new XYZ(point_stationCorner02.X - station_s_wall - station_c_wall / 2, point_stationCorner02.Y - station_s_wall - station_c_wall / 2, point_stationCorner02.Z), new XYZ(point_stationCorner04.X + station_s_wall + station_c_wall / 2, point_stationCorner04.Y - station_s_wall - station_c_wall / 2, point_stationCorner04.Z));
            c_wallEdgeLine = new List<Line>();
            c_wallEdgeLine.Add(Edge01_c_wall);
            c_wallEdgeLine.Add(Edge02_c_wall);
            c_wallEdgeLine.Add(Edge03_c_wall);
            c_wallEdgeLine.Add(Edge04_c_wall);

            using (Transaction ttt = new Transaction(doc, "Create Wall API"))
            {
                ttt.Start();

                List<string> li = new List<string>();//----------------------篩選已經存在的
                foreach (WallType wallType in collector)
                {
                    li.Add(wallType.Name);
                }

                if (li.Contains(NewWallTypeName_s_wall) == false)
                {
                    foreach (WallType wallType in collector)//----------------------新增結構牆類型
                    {
                        if (wallType.Kind.ToString() == "Basic")
                        {
                            if (wallType.Name == "通用 - 200mm")//用其他方式再篩一次?確保英文版適用
                            {
                                NewWallType_s_wall = wallType.Duplicate(NewWallTypeName_s_wall) as WallType;
                                CompoundStructure cs = NewWallType_s_wall.GetCompoundStructure();
                                IList<CompoundStructureLayer> IstLayers = cs.GetLayers();
                                foreach (CompoundStructureLayer item in IstLayers)
                                {
                                    if (item.Function == MaterialFunctionAssignment.Structure)
                                    {
                                        item.Width = Convert.ToDouble(JB_stationData["s_wall"].ToString()) / 0.3048;
                                        break;
                                    }
                                }
                                cs.SetLayers(IstLayers);
                                NewWallType_s_wall.SetCompoundStructure(cs);

                                break;
                            }
                        }
                    }
                }
                else if (li.Contains(NewWallTypeName_s_wall) == true)
                    {
                    foreach (WallType wallType in collector)
                    {
                        if (wallType.Name == NewWallTypeName_s_wall)
                        {
                            NewWallType_s_wall = wallType;
                        }
                    }
                }

                if (li.Contains(NewWallTypeName_c_wall) == false)
                {
                    foreach (WallType wallType in collector)//----------------------新增連續壁類型
                    {
                        if (wallType.Kind.ToString() == "Basic")
                        {
                            if (wallType.Name == "通用 - 200mm")//用其他方式再篩一次?確保英文版適用
                            {
                                NewWallType_c_wall = wallType.Duplicate(NewWallTypeName_c_wall) as WallType;
                                CompoundStructure cs2 = NewWallType_c_wall.GetCompoundStructure();
                                IList<CompoundStructureLayer> IstLayers2 = cs2.GetLayers();
                                foreach (CompoundStructureLayer item in IstLayers2)
                                {
                                    if (item.Function == MaterialFunctionAssignment.Structure)
                                    {
                                        item.Width = Convert.ToDouble(JB_stationData["c_wall"].ToString()) / 0.3048;
                                        break;
                                    }
                                }
                                cs2.SetLayers(IstLayers2);
                                NewWallType_c_wall.SetCompoundStructure(cs2);

                                break;
                            }
                        }
                    }
                }
                else if (li.Contains(NewWallTypeName_c_wall) == true)
                {
                    foreach (WallType wallType in collector)
                    {
                        if (wallType.Name == NewWallTypeName_c_wall)
                        {
                            NewWallType_c_wall = wallType;
                        }
                    }
                }

                collector = new FilteredElementCollector(doc);
                //ICollection<Element> newlevels = collector.OfClass(typeof(Level)).ToElements();
                IList<Element> newlevels = collector.OfClass(typeof(Level)).ToElements();
                double levelHeight = 0;
                List<double> levelHeigtList = new List<double>();
                List<ElementId> levelDataList = new List<ElementId>();
                List<double> wallPointListX = new List<double>();
                List<double> wallPointListY = new List<double>();
                List<Wall> s_wallList = new List<Wall>();

                for (int i = 0; i < newlevels.Count; i++)//----------------------製作結構牆
                {
                    if (i < newlevels.Count - 1)
                    {
                        levelHeight = Convert.ToDouble(JB_levelData[i + 1]["height"].ToString()) / 0.3048 - Convert.ToDouble(JB_levelData[i]["height"].ToString()) / 0.3048;
                        levelHeigtList.Add(levelHeight);
                    }
                    else
                    {
                        levelHeigtList.Add(0.1);
                    }
                }

                foreach (Element elm in newlevels)
                {
                    levelDataList.Add(elm.Id);
                }

                for (int i = 0; i < newlevels.Count; i++)
                {
                    wallPointListX = new List<double>();
                    wallPointListY = new List<double>();
                    s_wallList = new List<Wall>();

                    foreach (Line lin in s_wallEdgeLine)
                    {
                        Wall wall = Wall.Create(doc, lin, NewWallType_s_wall.Id, levelDataList[i], levelHeigtList[i], 0, false, false);

                        LocationCurve lc = wall.Location as LocationCurve;
                        XYZ PSt = lc.Curve.GetEndPoint(0);
                        XYZ PEn = lc.Curve.GetEndPoint(1);
                        XYZ PAv = new XYZ((PSt.X + PEn.X) / 2, (PSt.Y + PEn.Y) / 2, 0);//----------------------------------------------------------------取牆中點，改結構牆定位線
                        wallPointListX.Add(PAv.X);
                        wallPointListY.Add(PAv.Y);
                        s_wallList.Add(wall);
                    }

                    double MaxX = wallPointListX.Max();
                    double MinX = wallPointListX.Min();
                    double MaxY = wallPointListY.Max();
                    double MinY = wallPointListY.Min();

                    for (int k = 0; k < s_wallList.Count; k++)
                    {
                        if (wallPointListX[k] == MaxX)
                        {
                            Parameter pa = s_wallList[k].get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                            pa.Set(3);
                        }
                        else if (wallPointListX[k] == MinX)
                        {
                            Parameter pa = s_wallList[k].get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                            pa.Set(2);
                        }
                        else if (wallPointListY[k] == MaxY)
                        {
                            Parameter pa = s_wallList[k].get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                            pa.Set(3);
                        }
                        else if (wallPointListY[k] == MinY)
                        {
                            Parameter pa = s_wallList[k].get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                            pa.Set(2);
                        }
                    }
                }

                List<double> indexList = new List<double>();

                for (int i = 0; i < newlevels.Count; i++)//----------------------製作連續壁
                {
                    double indexNumber = Convert.ToDouble(JB_levelData[i]["height"].ToString());
                    indexList.Add(indexNumber);
                }
                double MaxIndex = indexList.Max();
                double MinIndex = indexList.Min();

                Level maxLevel = null;
                for (int i = 0; i < newlevels.Count; i++)
                {
                    if (Convert.ToDouble(JB_levelData[i]["height"].ToString()) == MinIndex)
                    {
                        maxLevel = newlevels[i] as Level;
                    }
                }

                wallPointListX = new List<double>();
                wallPointListY = new List<double>();
                s_wallList = new List<Wall>();

                foreach (Curve lin in c_wallEdgeLine)
                {
                    Wall wall = Wall.Create(doc, lin, NewWallType_c_wall.Id, maxLevel.Id, MaxIndex / 0.3048 - MinIndex / 0.3048, 0, false, false);
                    Parameter pa = wall.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);

                    LocationCurve lc = wall.Location as LocationCurve;
                    XYZ PSt = lc.Curve.GetEndPoint(0);
                    XYZ PEn = lc.Curve.GetEndPoint(1);
                    XYZ PAv = new XYZ((PSt.X + PEn.X) / 2, (PSt.Y + PEn.Y) / 2, 0);//----------------------------------------------------------------取牆中點，改連續壁定位線
                    wallPointListX.Add(PAv.X);
                    wallPointListY.Add(PAv.Y);
                    s_wallList.Add(wall);
                }

                double MaxX_ = wallPointListX.Max();
                double MinX_ = wallPointListX.Min();
                double MaxY_ = wallPointListY.Max();
                double MinY_ = wallPointListY.Min();

                for (int k = 0; k < s_wallList.Count; k++)
                {
                    if (wallPointListX[k] == MaxX_)
                    {
                        Parameter pa = s_wallList[k].get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                        pa.Set(2);
                    }
                    else if (wallPointListX[k] == MinX_)
                    {
                        Parameter pa = s_wallList[k].get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                        pa.Set(3);
                    }
                    else if (wallPointListY[k] == MaxY_)
                    {
                        Parameter pa = s_wallList[k].get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                        pa.Set(2);
                    }
                    else if (wallPointListY[k] == MinY_)
                    {
                        Parameter pa = s_wallList[k].get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM);
                        pa.Set(3);
                    }
                }

                    ttt.Commit();
            }

            FloorType NewFloorType = null;
            string NewFloorTypeName = "站體樓板";

            using (Transaction ttt = new Transaction(doc, "Create Floor API"))
            {
                ttt.Start();

                CurveArray profile = new CurveArray();
                profile.Append(Line.CreateBound(point_stationCorner01, point_stationCorner02));
                profile.Append(Line.CreateBound(point_stationCorner02, point_stationCorner04));
                profile.Append(Line.CreateBound(point_stationCorner04, point_stationCorner03));
                profile.Append(Line.CreateBound(point_stationCorner03, point_stationCorner01));

                collector = new FilteredElementCollector(doc).OfClass(typeof(FloorType));

                List<string> li = new List<string>();//----------------------篩選已經存在的
                foreach (FloorType floorType in collector)
                {
                    li.Add(floorType.Name);
                }

                if (li.Contains(NewFloorTypeName) == false)
                {
                    foreach (FloorType floorType in collector)
                    {
                        if (floorType.Name == "通用 300mm")
                        {
                            NewFloorType = floorType.Duplicate(NewFloorTypeName) as FloorType;
                            CompoundStructure cs = NewFloorType.GetCompoundStructure();
                            IList<CompoundStructureLayer> IstLayers = cs.GetLayers();
                            foreach (CompoundStructureLayer item in IstLayers)
                            {
                                if (item.Function == MaterialFunctionAssignment.Structure)
                                {
                                    item.Width = Convert.ToDouble(JB_levelData[0]["structure_THK"].ToString()) / 0.3048;
                                    break;
                                }
                            }
                            cs.SetLayers(IstLayers);
                            NewFloorType.SetCompoundStructure(cs);

                            break;
                        }
                    }
                }

                else if (li.Contains(NewFloorTypeName) == true)
                {
                    foreach (FloorType floorType in collector)
                    {
                        if (floorType.Name == NewFloorTypeName)
                        {
                            NewFloorType = floorType;
                        }
                    }
                }

                collector = new FilteredElementCollector(doc);
                ICollection<Element> newlevels = collector.OfClass(typeof(Level)).ToElements();//----------------------建立樓板
                XYZ normal = XYZ.BasisZ;

                foreach (Element elm in newlevels)
                {
                    //Floor floor = doc.Create.NewFloor(profile, NewFloorType, elm as Level, true, normal);
                    //Parameter pa = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                    //pa.Set(0);
                }

                ttt.Commit();
            }

            IList<Element> diglevels = collector.OfClass(typeof(Level)).ToElements();
            Level level_ = null;
            List<double> digsHeightList = new List<double>();
            List<Level> digsLevelList = new List<Level>();
            List<Line> dig_EdgeList = new List<Line>();

            using (Transaction ttt = new Transaction(doc, "Create Dig API"))//----------------------建立擴挖區
            {
                ttt.Start();

                for (int i = 0; i < JB_digData.Count; i++)
                {
                    for (int j = 0; j < JB_levelData.Count; j++)
                    {
                        if (JB_digData[i]["Level_id"].ToString() == JB_levelData[j]["index"].ToString())
                        {
                            double digHeight = Convert.ToDouble(JB_levelData[j + 1]["height"]) / 0.3048 - Convert.ToDouble(JB_levelData[j]["height"]) / 0.3048;
                            digsHeightList.Add(digHeight);
                            //MessageBox.Show(digHeight.ToString());
                        }
                    }
                }

                for (int i = 0; i < JB_digData.Count; i++)
                {
                    for (int j = 0; j < JB_levelData.Count; j++)
                    {
                        if (JB_digData[i]["Level_id"].ToString() == JB_levelData[j]["index"].ToString())
                        {
                            string levelNme = JB_levelData[j]["name"].ToString();

                            foreach(Element el in diglevels)
                            {
                                if(el.Name == levelNme)
                                {
                                    level_ = (Level)el;
                                    digsLevelList.Add(level_);
                                    //MessageBox.Show(level_.Name);
                                }
                            }
                        }
                    }
                }

                for(int i = 0; i < JB_digData.Count; i++)
                {
                    XYZ dig_DatumPoint = new XYZ(Math.Round(Convert.ToDouble(JB_digData[i]["Dig_x"]), 2) / 0.3048, Convert.ToDouble(JB_digData[i]["Dig_y"]) / 0.3048, 0);
                    XYZ dig_1QuadrantPoint = new XYZ(Math.Round(Convert.ToDouble(JB_digData[i]["Dig_x"]), 2) / 0.3048 + Convert.ToDouble(JB_digData[i]["Length"]) / 0.3048, Convert.ToDouble(JB_digData[i]["Dig_y"]) / 0.3048, 0);
                    XYZ dig_3QuadrantPoint = new XYZ(Math.Round(Convert.ToDouble(JB_digData[i]["Dig_x"]), 2) / 0.3048, Convert.ToDouble(JB_digData[i]["Dig_y"]) / 0.3048 - Convert.ToDouble(JB_digData[i]["Width"]) / 0.3048, 0);
                    XYZ dig_4QuadrantPoint = new XYZ(Math.Round(Convert.ToDouble(JB_digData[i]["Dig_x"]), 2) / 0.3048 + Convert.ToDouble(JB_digData[i]["Length"]) / 0.3048, Convert.ToDouble(JB_digData[i]["Dig_y"]) / 0.3048 - Convert.ToDouble(JB_digData[i]["Width"]) / 0.3048, 0);

                    Line dig_Edge01 = Line.CreateBound(dig_DatumPoint, dig_1QuadrantPoint);
                    Line dig_Edge02 = Line.CreateBound(dig_DatumPoint, dig_3QuadrantPoint);
                    Line dig_Edge03 = Line.CreateBound(dig_1QuadrantPoint, dig_4QuadrantPoint);
                    Line dig_Edge04 = Line.CreateBound(dig_3QuadrantPoint, dig_4QuadrantPoint);

                    dig_EdgeList = new List<Line>();
                    dig_EdgeList.Add(dig_Edge01);
                    dig_EdgeList.Add(dig_Edge02);
                    dig_EdgeList.Add(dig_Edge03);
                    dig_EdgeList.Add(dig_Edge04);

                    foreach(Line lin in dig_EdgeList)
                    {
                        Wall wall = Wall.Create(doc, lin, NewWallType_s_wall.Id, digsLevelList[i].Id, digsHeightList[i], 0, false, false);


                        //LocationCurve lc = wall.Location as LocationCurve;
                        //lc.Curve.GetEndPoint(0);
                    }
                }

                ttt.Commit();
            }


            return Result.Succeeded;

            /*  引入外部model直接讀取的方法
            list = JsonConvert.DeserializeObject<List<MRTStation>>(steee);
            foreach (MRTStation strr in list)
            {
                MessageBox.Show(strr.width.ToString());

            }*/

        }




    }

    
}
