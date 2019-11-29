Imports System.Diagnostics
Imports System.Collections
Imports System.Windows.Forms
Imports System.Drawing


Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Namespace A2acad
    Partial Public Class A2acad
        Public Sub CreateMyTable()
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = CType(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = CType(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                Dim tbl As Table = New Table()
                tbl.TableStyle = db.Tablestyle
                tbl.Position = ed.GetPoint(vbLf & "Pick a point: ").Value
                Dim ts As TableStyle = CType(tr.GetObject(tbl.TableStyle, OpenMode.ForRead), TableStyle)
                Dim textht As Double = ts.TextHeight(RowType.DataRow)
                Dim rows As Integer = 10
                Dim columns As Integer = 4
                tbl.InsertRows(1, textht * 2, rows)
                tbl.InsertColumns(1, textht * 15, columns)
                Dim range As CellRange = CellRange.Create(tbl, 0, 0, 0, columns)
                tbl.MergeCells(range)
                'tbl.Cells(0, 0).Style = "Titulo"    '"Title"
                tbl.Cells(0, 0).TextString = "Title"
                tbl.Rows(0).Height = textht * 2
                tbl.InsertRows(1, textht * 2, 1)
                tbl.Rows(1).Style = "Header"
                tbl.Rows(1).Height = textht * 1.5
                tbl.Cells(1, 0).Contents.Add()
                tbl.Cells(1, 0).Contents(0).TextString = "Header #1"
                For c As Integer = 1 To columns
                    tbl.Cells(1, c).TextString = "Header  #" & (c + 1).ToString()
                Next

                For r As Integer = 2 To rows + 2 - 1
                    'tbl.Rows(r).Style = "Data"
                    tbl.Rows(r).Height = textht * 1.25
                    tbl.Cells(r, 0).Contents.Add()
                    tbl.Cells(r, 0).Contents(0).TextString = "DataRow  #" & (r - 1).ToString() & " Col 1"
                    For c As Integer = 1 To columns
                        tbl.Cells(r, c).TextString = "DataRow  #" & (r - 1).ToString() & " Col " + (c + 1).ToString()
                    Next
                Next

                For Each col As Column In tbl.Columns
                    col.Width = textht * 15
                Next

                Dim dtp As DataTypeParameter = New DataTypeParameter()
                dtp.DataType = DataType.Double
                dtp.UnitType = UnitType.Distance
                For r As Integer = 2 To rows + 2 - 1
                    tbl.Cells(r, columns).Contents(0).DataType = dtp
                    tbl.Cells(r, columns).Contents(0).Value = Math.Pow(Math.PI, 1 / r)
                    tbl.Cells(r, columns).Contents(0).DataFormat = "%lu2%pr3%th44"
                Next

                tbl.GenerateLayout()
                btr.AppendEntity(tbl)
                tr.AddNewlyCreatedDBObject(tbl, True)
                tr.Commit()
            End Using
        End Sub



        '[CommandMethod("CreatePlainTable")]
        'Public void CreateMyTable()
        '    {
        '        // based on code written by Kean Walmsley
        '        Document doc = Application.DocumentManager.MdiActiveDocument;
        '        Database db = doc.Database;
        '        Editor ed = doc.Editor;
        '        Using (Transaction tr = db.TransactionManager.StartTransaction())
        '        {
        '            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
        '            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
        '            Table tbl = New Table();
        '            tbl.TableStyle = db.Tablestyle;
        '            tbl.Position = ed.GetPoint("\nPick a point: ").Value;
        '            TableStyle ts = (TableStyle)tr.GetObject(tbl.TableStyle, OpenMode.ForRead);
        '            Double textht = ts.TextHeight(RowType.DataRow);
        '            int rows = 10;
        '            int columns = 4;
        '            //insert rows
        '            tbl.InsertRows(1, textht * 2, rows);
        '            // insert columns
        '            tbl.InsertColumns(1, textht * 15, columns);// first column Is already exist, thus we'll have 5 columns
        '            //create range to merge the cells in the first row
        '            CellRange range = CellRange.Create(tbl, 0, 0, 0, columns);
        '            tbl.MergeCells(range);
        '            // set style for title row
        '            tbl.Cells[0, 0].Style = "Title";
        '            tbl.Cells[0, 0].TextString = "Title";

        '            tbl.Rows[0].Height = textht * 2;
        '            tbl.InsertRows(1, textht * 2, 1);
        '            // set style for header row
        '            tbl.Rows[1].Style = "Header";
        '            tbl.Rows[1].Height = textht * 1.5;
        '            //create contents in the first cell And set textstring
        '            tbl.Cells[1, 0].Contents.Add();
        '            tbl.Cells[1, 0].Contents[0].TextString = "Header #1";
        '            For (int c = 1; c <= columns; c++)
        '            {
        '                //for all of the rest cells just set textstring (Or value)
        '                tbl.Cells[1, c].TextString = "Header  #" + (c + 1).ToString();
        '            }

        '            For (int r = 2; r < rows + 2; r++)//exact number of data rows + title row + header row
        '            {
        '                // set style for data row
        '                tbl.Rows[r].Style = "Data";
        '                tbl.Rows[r].Height = textht * 1.25;
        '                //create contents in the first cell And set textstring
        '                tbl.Cells[r, 0].Contents.Add();
        '                tbl.Cells[r, 0].Contents[0].TextString = "DataRow  #" + (r - 1).ToString() + " Col 1";
        '                For (int c = 1; c <= columns; c++)
        '                {
        '                    //for all of the rest cells just set textstring (Or value)
        '                    tbl.Cells[r, c].TextString = "DataRow  #" + (r - 1).ToString() + " Col " + (c + 1).ToString(); ;
        '                }
        '            }
        '            // set equal column widths
        '            foreach (Column col in tbl.Columns)
        '                col.Width = textht * 15;

        '            //change last column values just to show data formatting               
        '            // to set numeric values with precision of 3 decimals:
        '            // create DataTypeParameter object
        '            // set data type,set value, then data format for every cell:

        '            DataTypeParameter dtp = New DataTypeParameter();
        '            dtp.DataType = DataType.Double;
        '            dtp.UnitType = UnitType.Distance;  // Or  UnitType.Unitless 

        '            //populate column with dummy values
        '            For (int r = 2; r < rows + 2; r++)//exact number of data rows + title row + header row  
        '            {
        '                tbl.Cells[r, columns].Contents[0].DataType = dtp;
        '                tbl.Cells[r, columns].Contents[0].Value = Math.Pow(Math.PI, 1.0 / r);
        '                tbl.Cells[r, columns].Contents[0].DataFormat = "%lu2%pr3%th44";//Or "%lu2%pr3%"
        '            }
        '            tbl.GenerateLayout();
        '            btr.AppendEntity(tbl);
        '            tr.AddNewlyCreatedDBObject(tbl, true);
        '            tr.Commit();
        '        }
        '    }
    End Class
End Namespace