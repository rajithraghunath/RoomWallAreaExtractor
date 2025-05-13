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
public class WallSplit : IExternalCommand
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

                            // Disallow joins
                            WallUtils.DisallowWallJoinAtEnd(wall, 0);
                            WallUtils.DisallowWallJoinAtEnd(wall, 1);

                            Curve curve = locCurve.Curve;
                            if (curve == null || curve.Length < 0.1)
                            {
                                skippedWalls.Add($"Skipped wall ID {wall.Id} due to short/invalid curve.");
                                continue;
                            }

                            // Orientation
                            XYZ direction = (curve.GetEndPoint(1) - curve.GetEndPoint(0)).Normalize();
                            string orientation = GetOrientation(direction);

                            // Area
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
            return dir.X > 0 ? "East Facing" : "West Facing";
        else
            return dir.Y > 0 ? "North Facing" : "South Facing";
    }
}
