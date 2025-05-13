using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

namespace RoomWallAreaExtractor
{
    [Transaction(TransactionMode.Manual)]
    public class ExtractWallArea : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active document.
            Document doc = commandData.Application.ActiveUIDocument.Document;

            // Collect all room elements.
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            // Prepare a list of CSV lines.
            // CSV header: RoomNumber, RoomName, WallId, WallArea, TotalRoomWallArea (for the room)
            List<string> csvLines = new List<string>
            {
                "Rm. No.,Room Name,Wall Id,Wall Area,Total Room Wall Area"
            };

            using (Transaction trans = new Transaction(doc, "Extract Wall Areas"))
            {
                trans.Start();

                // Loop through each room.
                foreach (Element roomElement in roomCollector)
                {
                    Room room = roomElement as Room;
                    if (room != null)
                    {
                        string roomName = room.Name;
                        string roomNumber = room.Number;

                        // Modify the room name to exclude the room number
                        if (roomName.EndsWith(" " + roomNumber))
                        {
                            roomName = roomName.Substring(0, roomName.Length - roomNumber.Length - 1);
                            room.Name = roomName; // Update the room name
                        }



                        double roomWallArea = 0.0;
                        // List to store individual wall info (Wall Id, Wall Area)
                        List<Tuple<string, double>> wallDetails = new List<Tuple<string, double>>();

                        // Retrieve the room boundary segments.
                        SpatialElementBoundaryOptions options = new SpatialElementBoundaryOptions();
                        IList<IList<BoundarySegment>> boundaries = room.GetBoundarySegments(options);

                        if (boundaries != null)
                        {
                            // Use a set to avoid processing duplicate walls.
                            HashSet<ElementId> processedWalls = new HashSet<ElementId>();
                            foreach (IList<BoundarySegment> boundaryList in boundaries)
                            {
                                foreach (BoundarySegment segment in boundaryList)
                                {
                                    ElementId wallId = segment.ElementId;
                                    if (wallId != ElementId.InvalidElementId && !processedWalls.Contains(wallId))
                                    {
                                        processedWalls.Add(wallId);
                                        Element wallElement = doc.GetElement(wallId);
                                        Autodesk.Revit.DB.Wall wall = wallElement as Autodesk.Revit.DB.Wall;
                                        if (wall != null)
                                        {
                                            // Retrieve the computed wall area.
                                            Parameter areaParam = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                                            if (areaParam != null)
                                            {
                                                double area = areaParam.AsDouble();
                                                roomWallArea += area;
                                                wallDetails.Add(new Tuple<string, double>(wall.Id.IntegerValue.ToString(), area));
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // For each wall in the room, add a CSV row.
                        // Format: RoomNumber, RoomName, WallId, WallArea, TotalRoomWallArea
                        if (wallDetails.Count > 0)
                        {
                            foreach (var wallInfo in wallDetails)
                            {
                                string line = string.Format("{0},{1},{2},{3:F2},{4:F2}",
                                    roomNumber,
                                    roomName,
                                    wallInfo.Item1,
                                    wallInfo.Item2,
                                    roomWallArea);
                                csvLines.Add(line);
                            }
                        }
                        else
                        {
                            // In case no wall details were found, output the room info.
                            string line = string.Format("{0},{1},,0.00,0.00", roomNumber, roomName);
                            csvLines.Add(line);
                        }
                    }
                }

                trans.Commit();
            }

            // Define file path on the Desktop.
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktopPath, "RoomWallAreas.csv");

            // Write all CSV lines to the file.
            File.WriteAllLines(filePath, csvLines);

            // Inform the user where the CSV file has been saved.
            TaskDialog.Show("Export Complete", "CSV file exported to:\n" + filePath);
            return Result.Succeeded;
        }
    }
}
