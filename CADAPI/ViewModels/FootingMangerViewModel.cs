using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using CADAPI.Commands;
using CADAPI.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using AutoCADApp = Autodesk.AutoCAD.ApplicationServices.Application;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using Line = Autodesk.AutoCAD.DatabaseServices.Line;


namespace CADAPI.ViewModels
{
    public class FootingMangerViewModel : INotifyPropertyChanged
    {

        #region Feilds
        private string _excelPath;
        private bool _drawTags;
        private bool _showTable;
        #endregion

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;


        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region Constructor
        public FootingMangerViewModel()
        {
            Browse = new Command(_ => BrowseExcel());
            DrawFooting = new Command(_ => ImportFootingsFromExcel(), _ => !string.IsNullOrWhiteSpace(ExcelPath));
        }


        #endregion

        #region Properities
        public Command Browse { get; }
        public Command DrawFooting { get; }

        public string ExcelPath
        {
            get => _excelPath;
            set { _excelPath = value; OnPropertyChanged(nameof(ExcelPath)); }
        }
        public bool DrawTags
        {
            get => _drawTags;
            set
            {
                _drawTags = value;
                OnPropertyChanged(nameof(DrawTags));
            }
        }
        public bool ShowTable
        {
            get => _showTable;
            set
            {
                _showTable = value;
                OnPropertyChanged(nameof(ShowTable));
            }
        }

        public System.Action CloseAction { get; internal set; }



        #endregion

        #region Methods
        private string BrowseExcel()
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls",
                Title = "Select Excel File"
            };

            if (dialog.ShowDialog() == true)
            {
                ExcelPath = dialog.FileName;
                return ExcelPath;
            }
            return string.Empty;
        }
        public void ImportFootingsFromExcel()
        {
            string filePath = ExcelPath;

            List<FootingData> footings = ReadFootingsFromExcel(filePath);
            if (footings.Count == 0)
            {
                AutoCADApp.ShowAlertDialog("No footing data found.");
                return;
            }

            DrawFootings(footings);
            AutoCADApp.ShowAlertDialog($"{footings.Count} footings imported successfully.");
            //AutoCADApp.ShowAlertDialog("Drawing completed!");
            if (ShowTable)
            {
                DrawFootingSectionSample();
                // InsertImage();
                DrawExcelTableInAutoCAD(filePath);
            }

            CloseAction?.Invoke();
        }
        private List<FootingData> ReadFootingsFromExcel(string filePath)

        {
            List<FootingData> footings = new List<FootingData>();
            ExcelApp excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel._Worksheet sheet = workbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range range = sheet.UsedRange;

                int row = 2;
                while (true)
                {
                    var xVal = range.Cells[row, 2].Value2;
                    var yVal = range.Cells[row, 3].Value2;
                    var widthPC = range.Cells[row, 4].Value2;
                    var lengthPC = range.Cells[row, 5].Value2;
                    var depthPC = range.Cells[row, 6].Value2;
                    var widthRC = range.Cells[row, 7].Value2;
                    var lengthRC = range.Cells[row, 8].Value2;
                    var depthRC = range.Cells[row, 9].Value2;


                    if (xVal == null && yVal == null && widthPC == null &&
                        lengthPC == null && depthPC == null &&
                        widthRC == null && lengthRC == null && depthRC == null)
                        break;

                    try
                    {
                        footings.Add(new FootingData
                        {
                            X = Convert.ToDouble(xVal),
                            Y = Convert.ToDouble(yVal),
                            WidthPC = Convert.ToDouble(widthPC),
                            LengthPC = Convert.ToDouble(lengthPC),
                            DepthPC = Convert.ToDouble(depthPC),
                            WidthRC = Convert.ToDouble(widthRC),
                            LengthRC = Convert.ToDouble(lengthRC),
                            DepthRC = Convert.ToDouble(depthRC)

                        });
                    }
                    catch (System.Exception ex)
                    {
                        AutoCADApp.ShowAlertDialog($"Error reading row {row}: {ex.Message}");
                    }

                    row++;
                }

                workbook.Close(false);
                excelApp.Quit();
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (System.Exception ex)
            {
                AutoCADApp.ShowAlertDialog($"Failed to read Excel file: {ex.Message}");
            }

            return footings;
        }

        public void DrawFootings(List<FootingData> footings)
        {
            Document doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                int index = 1; // Start numbering from F1

                foreach (var footing in footings)
                {
                    DrawRectangle(btr, footing.X, footing.Y, footing.WidthPC, footing.LengthPC, 7);
                    DrawRectangle(btr, footing.X, footing.Y, footing.WidthRC, footing.LengthRC, 7);

                    if (DrawTags)
                    {
                        string tag = $"F{index}";
                        DrawFootingTag(btr, footing.X, footing.Y, tag);
                        index++; // Move to next number
                    }
                }

                tr.Commit();
            }
        }


        private void DrawRectangle(BlockTableRecord btr, double centerX, double centerY, double width, double length, short colorIndex)
        {
            double halfW = width / 2;
            double halfL = length / 2;

            Point2d[] corners = new Point2d[]
            {
                new Point2d(centerX - halfW, centerY - halfL),
                new Point2d(centerX + halfW, centerY - halfL),
                new Point2d(centerX + halfW, centerY + halfL),
                new Point2d(centerX - halfW, centerY + halfL),
            };

            Polyline pline = new Polyline();
            for (int i = 0; i < 4; i++)
                pline.AddVertexAt(i, corners[i], 0, 0, 0);
            pline.Closed = true;
            pline.ColorIndex = colorIndex;

            btr.AppendEntity(pline);
            btr.Database.TransactionManager.TopTransaction.AddNewlyCreatedDBObject(pline, true);
        }
        private void DrawFootingTag(BlockTableRecord btr, double centerX, double centerY, string tag)
        {
            // Tag position offset
            double offsetX = 0.75;
            double offsetY = 0.75;
            double textHeight = 0.25;
            Point3d tagPosition = new Point3d(centerX + offsetX, centerY + offsetY, 0);

            // 1. Create DBText
            DBText text = new DBText
            {
                TextString = tag,
                Height = textHeight,
                Position = tagPosition,
                HorizontalMode = TextHorizontalMode.TextCenter,
                VerticalMode = TextVerticalMode.TextVerticalMid,
                AlignmentPoint = tagPosition
            };

            btr.AppendEntity(text);
            btr.Database.TransactionManager.TopTransaction.AddNewlyCreatedDBObject(text, true);

            // 2. Create surrounding circle
            double radius = textHeight * 1.2; // Slightly larger than text
            Circle circle = new Circle(tagPosition, Vector3d.ZAxis, radius);
            circle.ColorIndex = 2; // Optional: red

            btr.AppendEntity(circle);
            btr.Database.TransactionManager.TopTransaction.AddNewlyCreatedDBObject(circle, true);
        }



        public void DrawExcelTableInAutoCAD(string filePath)
        {
            Document doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ExcelApp excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;

            try
            {
                excelApp = new ExcelApp();
                workbook = excelApp.Workbooks.Open(filePath);
                Microsoft.Office.Interop.Excel._Worksheet sheet = workbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range usedRange = sheet.UsedRange;

                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                // Ask user to pick insertion point
                PromptPointResult ppr = ed.GetPoint("\nSelect insertion point for Excel table: ");
                if (ppr.Status != PromptStatus.OK) return;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    Table table = new Table();
                    table.TableStyle = db.Tablestyle;
                    table.SetSize(rowCount, colCount);
                    table.SetRowHeight(1.5);        // Compact height
                    table.SetColumnWidth(10);       // Compact width
                    table.Position = ppr.Value;     // User-defined location

                    // General cell formatting
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            Cell cell = table.Cells[r, c];
                            cell.TextHeight = 0.75;
                            cell.Alignment = CellAlignment.MiddleCenter;
                        }
                    }

                    // Fill cells from Excel data
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellVal = usedRange.Cells[row, col].Value2;
                            string cellText = cellVal != null ? cellVal.ToString() : " ";

                            Cell cell = table.Cells[row - 1, col - 1];
                            cell.TextString = string.IsNullOrWhiteSpace(cellText) ? " " : cellText;

                            // Style header row
                            if (row == 1)
                            {
                                cell.TextHeight = 1.0;
                                cell.Alignment = CellAlignment.MiddleCenter;
                                cell.BackgroundColor = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 209, 80); // #FFD150
                                cell.Borders.Top.LineWeight = LineWeight.LineWeight050;
                                cell.Borders.Bottom.LineWeight = LineWeight.LineWeight050;
                            }
                        }
                    }

                    btr.AppendEntity(table);
                    tr.AddNewlyCreatedDBObject(table, true);
                    tr.Commit();
                }

                workbook.Close(false);
                excelApp.Quit();
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
            catch (System.Exception ex)
            {
                AutoCADApp.ShowAlertDialog($"Failed to draw table: {ex.Message}");
            }
        }

        public void InsertImage()
        {
            // ✅ Step 1: Hardcoded image path
            string imagePath = @"D:\ITI\AutoCAD API\Final Project\Final\CADAPI\Images\xxxxcc.png"; // ← Change this to your actual image path

            var doc = AutoCADApp.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            if (!System.IO.File.Exists(imagePath))
            {
                ed.WriteMessage($"\nImage not found: {imagePath}");
                return;
            }

            // ✅ Step 2: Let user pick insertion point
            var ppr = ed.GetPoint("\nClick to insert image: ");
            if (ppr.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // ✅ Step 3: Ensure image dictionary exists
                ObjectId dictId = RasterImageDef.GetImageDictionary(db);
                if (dictId == ObjectId.Null)
                {
                    RasterImageDef.CreateImageDictionary(db);
                    dictId = RasterImageDef.GetImageDictionary(db);
                }

                DBDictionary imgDict = (DBDictionary)tr.GetObject(dictId, OpenMode.ForWrite);
                string imageName = System.IO.Path.GetFileNameWithoutExtension(imagePath);

                // ✅ Step 4: Load or reuse image definition
                RasterImageDef imgDef;
                if (!imgDict.Contains(imageName))
                {
                    imgDef = new RasterImageDef { SourceFileName = imagePath };
                    imgDef.Load(); // 🔄 Required
                    imgDict.SetAt(imageName, imgDef);
                    tr.AddNewlyCreatedDBObject(imgDef, true);
                }
                else
                {
                    imgDef = (RasterImageDef)tr.GetObject(imgDict.GetAt(imageName), OpenMode.ForRead);
                }

                // ✅ Step 5: Create RasterImage
                RasterImage image = new RasterImage();
                image.ImageDefId = imgDef.ObjectId;

                image.Orientation = new CoordinateSystem3d(
                    ppr.Value,                       // Position
                    new Vector3d(10, 0, 0),          // X scale vector
                    new Vector3d(0, 10, 0)           // Y scale vector
                );
                image.ShowImage = true;
                RasterImage.EnableReactors(true);

                // ✅ Step 6: Add image to ModelSpace, THEN associate
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                btr.AppendEntity(image);
                tr.AddNewlyCreatedDBObject(image, true);

                image.AssociateRasterDef(imgDef); // ✅ Associate only after adding to database

                tr.Commit();
            }
        }

        public void DrawFootingSectionSample()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptPointResult ppr = ed.GetPoint("\nSpecify insertion point for footing section:");
            if (ppr.Status != PromptStatus.OK) return;

            Point3d basePoint = ppr.Value;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    double footingWidth = 4.0;
                    double footingDepth = 0.6;
                    double columnHeight = 1.5;
                    double cover = 0.07;
                    double colCover = 0.025;
                    double rebarRadius = 0.008;
                    double spacing = 0.2;

                    Func<double, double, Point3d> pt = (x, y) => new Point3d(basePoint.X + x, basePoint.Y + y, 0);

                    // Footing and column outlines
                    var lines = new[]
                    {
                new Line(pt(0,0), pt(footingWidth,0)),
                new Line(pt(footingWidth,0), pt(footingWidth,footingDepth)),
                new Line(pt(footingWidth,footingDepth), pt(footingWidth/2+0.5,footingDepth)),
                new Line(pt(footingWidth/2-0.5,footingDepth), pt(0,footingDepth)),
                new Line(pt(0,footingDepth), pt(0,0)),
                new Line(pt(footingWidth/2+0.5,footingDepth), pt(footingWidth/2+0.5,footingDepth+columnHeight)),
                new Line(pt(footingWidth/2-0.5,footingDepth), pt(footingWidth/2-0.5,footingDepth+columnHeight))
            };

                    foreach (var l in lines)
                    {
                        btr.AppendEntity(l);
                        tr.AddNewlyCreatedDBObject(l, true);
                    }

                    // Short direction rebars
                    var rebars = new[]
                    {
                new Line(pt(cover * 2, footingDepth - cover / 2), pt(cover, footingDepth - cover)),
                new Line(pt(cover, footingDepth - cover), pt(cover, cover)),
                new Line(pt(cover, cover), pt(footingWidth - cover, cover)),
                new Line(pt(footingWidth - cover, cover), pt(footingWidth - cover, footingDepth - cover)),
                new Line(pt(footingWidth - cover, footingDepth - cover), pt(footingWidth - cover * 2, footingDepth - cover / 2)),
                new Line(pt(cover, footingDepth - cover), pt(footingWidth - cover, footingDepth - cover))
            };

                    foreach (var l in rebars)
                    {
                        btr.AppendEntity(l);
                        tr.AddNewlyCreatedDBObject(l, true);
                    }

                    // Long direction circular rebars
                    for (double x = cover + rebarRadius; x <= footingWidth - cover - rebarRadius; x += spacing)
                    {
                        var c = new Circle(pt(x, cover + rebarRadius), Vector3d.ZAxis, rebarRadius);
                        c.LineWeight = LineWeight.LineWeight100; 
                        c.ColorIndex = 1; // Red color
                        btr.AppendEntity(c);
                        tr.AddNewlyCreatedDBObject(c, true);
                    }

                    // Column rebars
                    var colRebars = new[]
                    {
                new Line(pt(footingWidth / 2 - 0.5 + colCover, cover), pt(footingWidth / 2 - 0.5 + colCover, footingDepth + columnHeight)),
                new Line(pt(footingWidth / 2 + 0.5 - colCover, cover), pt(footingWidth / 2 + 0.5 - colCover, footingDepth + columnHeight))
            };

                    foreach (var l in colRebars)
                    {
                        btr.AppendEntity(l);
                        tr.AddNewlyCreatedDBObject(l, true);
                    }

                    // === Dimension Indicators Only ===

                    var dimWidth = new RotatedDimension
                    {
                        XLine1Point = pt(0, 0),
                        XLine2Point = pt(footingWidth, 0),
                        DimLinePoint = pt(footingWidth / 2, -0.4),
                        Rotation = 0,
                        DimensionText = "B",
                        Dimltype = db.Dimstyle
                    };

                    var dimHeight = new RotatedDimension
                    {
                        XLine1Point = pt(footingWidth, 0),
                        XLine2Point = pt(footingWidth, footingDepth),
                        DimLinePoint = pt(footingWidth + 0.4, footingDepth / 2),
                        Rotation = Math.PI / 2,
                        DimensionText = "H",
                        Dimltype = db.Dimstyle
                    };

                    btr.AppendEntity(dimWidth);
                    tr.AddNewlyCreatedDBObject(dimWidth, true);

                    btr.AppendEntity(dimHeight);
                    tr.AddNewlyCreatedDBObject(dimHeight, true);

                    tr.Commit();

                    ed.Command("._ZOOM", "_E");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n❌ Error: {ex.Message}");
                }
            }
        }



        #endregion
    }
}
