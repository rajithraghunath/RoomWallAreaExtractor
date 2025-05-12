using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Architecture;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

[Transaction(TransactionMode.Manual)]
public class ExtractWallArea : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiApp = commandData.Application;
        UIDocument uidoc = uiApp.ActiveUIDocument;
        Document doc = uidoc.Document;

        string csvPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\RoomWallAreas.csv";

        var roomCollector = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType();

        List<string> csvLines = new List<string> {
            "Room Number,Room Name,Wall Id,Wall Length (ft),Wall Area (sqft),Wall Orientation"
        };

        List<string> skippedWalls = new List<string>();

        using (Transaction tx = new Transaction(doc, "Split and Export Room Wall Data"))
        {
            tx.Start();

            // Collect all walls for intersection checks
            var wallCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(w => w.Location is LocationCurve)
                .ToList();

            List<ElementId> originalWallsToDelete = new List<ElementId>();
            List<Wall> newWalls = new List<Wall>();

            foreach (Wall wall in wallCollector)
            {
                try
                {
                    WallUtils.DisallowWallJoinAtEnd(wall, 0);
                    WallUtils.DisallowWallJoinAtEnd(wall, 1);

                    LocationCurve locationCurve = wall.Location as LocationCurve;
                    if (locationCurve == null)
                        continue;

                    Curve curve = locationCurve.Curve;
                    XYZ start = curve.GetEndPoint(0);
                    XYZ end = curve.GetEndPoint(1);

                    List<XYZ> intersectionPoints = new List<XYZ>();

                    foreach (Wall otherWall in wallCollector)
                    {
                        if (otherWall.Id == wall.Id) continue;

                        LocationCurve otherLoc = otherWall.Location as LocationCurve;
                        if (otherLoc == null) continue;

                        Curve otherCurve = otherLoc.Curve;
                        SetComparisonResult result = curve.Intersect(otherCurve, out IntersectionResultArray results);

                        if (result == SetComparisonResult.Overlap && results != null)
                        {
                            foreach (IntersectionResult ir in results)
                            {
                                XYZ point = ir.XYZPoint;
                                if (point != null && curve.Project(point) != null)
                                    intersectionPoints.Add(point);
                            }
                        }
                    }

                    if (intersectionPoints.Count == 0)
                        continue;

                    intersectionPoints = intersectionPoints
                        .Distinct(new XYZComparer())
                        .OrderBy(p => curve.Project(p).Parameter)
                        .ToList();

                    List<Curve> splitSegments = new List<Curve>();
                    XYZ currentStart = start;

                    foreach (XYZ pt in intersectionPoints)
                    {
                        if (!pt.IsAlmostEqualTo(currentStart))
                        {
                            splitSegments.Add(Line.CreateBound(currentStart, pt));
                        }
                        currentStart = pt;
                    }

                    if (!currentStart.IsAlmostEqualTo(end))
                    {
                        splitSegments.Add(Line.CreateBound(currentStart, end));
                    }

                    foreach (Curve segment in splitSegments)
                    {
                        Wall newWall = Wall.Create(doc, segment, wall.LevelId, false);

                        // Find and set the Top Constraint
                        Level baseLevel = doc.GetElement(wall.LevelId) as Level;
                        Level upperLevel = GetImmediateUpperLevel(doc, baseLevel);

                        if (upperLevel != null)
                        {
                            Parameter topConstraintParam = newWall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE);
                            if (topConstraintParam != null && !topConstraintParam.IsReadOnly)
                            {
                                topConstraintParam.Set(upperLevel.Id);
                            }
                        }
                        else
                        {
                            // Optional: fallback to a fixed height if no upper level is found
                            Parameter unconnectedHeight = newWall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                            if (unconnectedHeight != null && !unconnectedHeight.IsReadOnly)
                            {
                                unconnectedHeight.Set(10.0); // 10 ft or as needed
                            }
                        }

                        newWalls.Add(newWall);

                    }

                    originalWallsToDelete.Add(wall.Id);
                }
                catch (Exception ex)
                {
                    skippedWalls.Add($"Wall ID {wall.Id}: {ex.Message}");
                }
            }

            foreach (ElementId id in originalWallsToDelete)
            {
                doc.Delete(id);
            }

            var updatedWallCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>()
                .Where(w => w.Location is LocationCurve)
                .ToList();

            foreach (Room room in roomCollector)
            {
                if (room.Location is LocationPoint locationPoint && room.Area > 0)
                {
                    SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                    IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);

                    if (boundaries == null) continue;

                    foreach (var boundaryList in boundaries)
                    {
                        foreach (BoundarySegment seg in boundaryList)
                        {
                            Element wallElem = doc.GetElement(seg.ElementId);
                            Wall wall = wallElem as Wall;

                            if (wall == null || !(wall.Location is LocationCurve locCurve))
                                continue;

                            Curve curve = locCurve.Curve;
                            if (curve == null || curve.Length < 0.1)
                            {
                                skippedWalls.Add($"Skipped wall ID {wall.Id} due to short/invalid curve.");
                                continue;
                            }

                            XYZ direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize();
                            string orientation = GetOrientation(direction);

                            double height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 10.0;
                            double area = UnitUtils.Convert(curve.Length * height, UnitTypeId.SquareFeet, UnitTypeId.SquareFeet);

                            csvLines.Add($"{room.Number},{room.Name},{wall.Id},{curve.Length:F2},{area:F2},{orientation}");
                        }
                    }
                }
            }

            tx.Commit();
        }

        File.WriteAllLines(csvPath, csvLines);

        if (skippedWalls.Any())
        {
            TaskDialog.Show("Skipped Walls", string.Join(Environment.NewLine, skippedWalls));
        }

        TaskDialog.Show("Export Complete", $"Room wall data exported to:\n{csvPath}");

        return Result.Succeeded;
    }

    private string GetOrientation(XYZ dir)
    {
        if (Math.Abs(dir.X) > Math.Abs(dir.Y))
            return dir.X > 0 && dir.Y < 0 ? "South Wall" : "North Wall";
        else
            return dir.Y > 0 && dir.X < 0 ? "East Wall" : "West Wall";
    }

    private Level GetImmediateUpperLevel(Document doc, Level baseLevel)
    {
        var levels = new FilteredElementCollector(doc)
            .OfClass(typeof(Level))
            .Cast<Level>()
            .Where(l => l.Elevation > baseLevel.Elevation)
            .OrderBy(l => l.Elevation)
            .ToList();

        return levels.FirstOrDefault(); // returns the closest level above
    }


    private class XYZComparer : IEqualityComparer<XYZ>
    {
        public bool Equals(XYZ x, XYZ y)
        {
            return x.IsAlmostEqualTo(y);
        }

        public int GetHashCode(XYZ obj)
        {
            return obj.GetHashCode();
        }
    }
}
