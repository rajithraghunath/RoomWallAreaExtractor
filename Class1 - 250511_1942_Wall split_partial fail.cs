using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

[Transaction(TransactionMode.Manual)]
public class ExtractWallArea : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIDocument uidoc = commandData.Application.ActiveUIDocument;
        Document doc = uidoc.Document;

        Dictionary<ElementId, List<XYZ>> wallSplitPoints = new Dictionary<ElementId, List<XYZ>>();
        Dictionary<ElementId, (Wall wall, Line line)> wallLineDict = new Dictionary<ElementId, (Wall, Line)>();

        FilteredElementCollector wallCollector = new FilteredElementCollector(doc).OfClass(typeof(Wall)).WhereElementIsNotElementType();

        foreach (Wall wall in wallCollector)
        {
            if (wall.Location is LocationCurve loc && loc.Curve is Line line)
            {
                wallLineDict[wall.Id] = (wall, line);
            }
        }

        // 1. Split walls at intersections with other walls
        foreach (var wall1 in wallLineDict)
        {
            Line curve1 = wall1.Value.line;

            foreach (var wall2 in wallLineDict)
            {
                if (wall1.Key == wall2.Key) continue;
                Line curve2 = wall2.Value.line;

                if (curve1.Intersect(curve2, out IntersectionResultArray results) == SetComparisonResult.Overlap && results != null)
                {
                    foreach (IntersectionResult result in results)
                    {
                        XYZ point = result.XYZPoint;
                        if (IsPointOnCurve(curve1, point))
                        {
                            if (!wallSplitPoints.ContainsKey(wall1.Key))
                                wallSplitPoints[wall1.Key] = new List<XYZ>();

                            if (!wallSplitPoints[wall1.Key].Exists(p => p.IsAlmostEqualTo(point)))
                                wallSplitPoints[wall1.Key].Add(point);
                        }
                    }
                }
            }
        }

        // 2. Split walls at room boundaries
        FilteredElementCollector roomCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType();

        foreach (Room room in roomCollector)
        {
            SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
            IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);
            if (boundaries == null) continue;

            foreach (IList<BoundarySegment> segmentList in boundaries)
            {
                foreach (BoundarySegment segment in segmentList)
                {
                    ElementId wallId = segment.ElementId;
                    if (wallLineDict.ContainsKey(wallId))
                    {
                        XYZ pt = segment.GetCurve().GetEndPoint(0);
                        if (!wallSplitPoints.ContainsKey(wallId))
                            wallSplitPoints[wallId] = new List<XYZ>();

                        if (!wallSplitPoints[wallId].Exists(p => p.IsAlmostEqualTo(pt)))
                            wallSplitPoints[wallId].Add(pt);
                    }
                }
            }
        }

        // 3. Split Walls
        Transaction trans = new Transaction(doc, "Split Walls");
        trans.Start();

        List<Wall> newWalls = new List<Wall>();
        string skippedSegmentsLog = "";

        foreach (var entry in wallSplitPoints)
        {
            Wall originalWall = wallLineDict[entry.Key].wall;
            Line baseLine = wallLineDict[entry.Key].line;
            List<XYZ> points = entry.Value.OrderBy(p => baseLine.Project(p).Parameter).ToList();

            if (points.Count < 1) continue;

            XYZ start = baseLine.GetEndPoint(0);
            XYZ end = baseLine.GetEndPoint(1);

            List<XYZ> splitPoints = new List<XYZ> { start };
            splitPoints.AddRange(points);
            splitPoints.Add(end);

            WallType wallType = originalWall.WallType;
            ElementId levelId = originalWall.LevelId;
            double height = originalWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();

            for (int i = 0; i < splitPoints.Count - 1; i++)
            {
                XYZ p1 = splitPoints[i];
                XYZ p2 = splitPoints[i + 1];

                if (p1.IsAlmostEqualTo(p2) || p1.DistanceTo(p2) < 0.01)
                    continue;

                try
                {
                    Line segment = Line.CreateBound(p1, p2);

                    // Skip vertical segments
                    XYZ dir = (p2 - p1).Normalize();
                    if (Math.Abs(dir.X) < 0.001 && Math.Abs(dir.Y) < 0.001)
                    {
                        skippedSegmentsLog += $"Vertical segment skipped: Start: {p1}, End: {p2}\n";
                        continue;
                    }

                    if (segment == null || segment.Length < 0.01)
                        continue;

                    Wall newWall = Wall.Create(doc, segment, wallType.Id, levelId, height, 0, false, false);
                    newWalls.Add(newWall);
                }
                catch (Exception ex)
                {
                    string error = $"Start: {p1}, End: {p2}\n{ex.Message}";
                    TaskDialog.Show("RoomWallArea - Wall Segment Skipped", error);
                    skippedSegmentsLog += error + "\n";
                    continue;
                }
            }

            doc.Delete(originalWall.Id);
        }

        trans.Commit();

        // 4. Extract room-wall areas to CSV
        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string filePath = Path.Combine(desktop, "RoomWallAreas.csv");

        using (StreamWriter writer = new StreamWriter(filePath))
        {
            writer.WriteLine("RoomName,RoomNumber,WallId,WallArea_SqFt");

            foreach (Room room in roomCollector)
            {
                SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);
                if (boundaries == null) continue;

                foreach (IList<BoundarySegment> segmentList in boundaries)
                {
                    foreach (BoundarySegment segment in segmentList)
                    {
                        ElementId wallId = segment.ElementId;
                        Wall wall = doc.GetElement(wallId) as Wall;
                        if (wall == null) continue;

                        double area = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED).AsDouble();
                        double areaSqFt = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareFeet);

                        writer.WriteLine($"{room.Name},{room.Number},{wall.Id},{areaSqFt:F2}");
                    }
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(skippedSegmentsLog))
        {
            File.WriteAllText(Path.Combine(desktop, "SkippedWallSegments.log"), skippedSegmentsLog);
        }

        TaskDialog.Show("Complete", "Walls split and area data exported to Desktop as RoomWallAreas.csv\nSee SkippedWallSegments.log for skipped segments.");
        return Result.Succeeded;
    }

    private bool IsPointOnCurve(Line line, XYZ point, double tolerance = 0.01)
    {
        double param = line.Project(point).Parameter;
        return param > tolerance && param < (1 - tolerance);
    }
}
