using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;

namespace LTaskTestApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            

        }
        string partFilePath = "D:\\onedrive shyam\\LTask\\LPart.ipt";
        string DrawingSheetFilePath = "D:\\onedrive shyam\\LTask\\StandardSheet.idw";
        private void btn_ClickMe_Click(object sender, EventArgs e)
        {
            // INV INI
            Inventor.Application invapp = null; try
            {
                // Attempt to get a reference to a running instance of Inventor.
                invapp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");

            }
            catch
            {
                // Start Inventor.
                System.Type oType = System.Type.GetTypeFromProgID("Inventor.Application");
                invapp = (Inventor.Application)System.Activator.CreateInstance(oType);

                while (!invapp.Ready)
                {

                }

                // Make Inventor visible.
                invapp.Visible = true;
            }

            //PART INI
            PartDocument partDoc = null;
            partDoc = (PartDocument)invapp.Documents.Open(partFilePath, false);
            // DRW INI
            DrawingDocument drawingDoc = null;
            drawingDoc = (DrawingDocument)invapp.Documents.Open(DrawingSheetFilePath, true);
            //PART CALL, VISIBLE OFF

            //DRW TEMPLATE CALL

            //ADD BASEVIEW
            Sheet oSheet = null;
            oSheet = drawingDoc.ActiveSheet;

            Point2d point1 = null;
            point1 = invapp.TransientGeometry.CreatePoint2d(8, 8);
            DrawingView drawingView = null;
            drawingView = oSheet.DrawingViews.AddBaseView((_Document)partDoc, point1, 0.015, ViewOrientationTypeEnum.kIsoTopRightViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle);
            drawingView.Label.FormattedText = "ISO-VIEW";

            DrawingView drawingView1 = null;
            point1 = invapp.TransientGeometry.CreatePoint2d(15, 20);
            drawingView1 = oSheet.DrawingViews.AddBaseView((_Document)partDoc, point1, 0.025, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle);
            drawingView1.Label.FormattedText = "FRONT-VIEW";


            //ADD DIMENSIONS

            Debug.WriteLine(drawingView1.DrawingCurves.Count);
            double topPoint = drawingView1.Position.Y + drawingView1.Height / 2;
            double BotPoint = drawingView1.Position.Y - drawingView1.Height / 2;

            double leftPoint = drawingView1.Position.X - drawingView1.Width / 2;
            double rightPoint = drawingView1.Position.X + drawingView1.Width / 2;

            double offset = (topPoint - drawingView1.Position.Y) / drawingView1.Scale;
            double offset_Bot = (BotPoint - drawingView1.Position.Y) / drawingView1.Scale;

            double offsetleft = (leftPoint - drawingView1.Position.X) / drawingView1.Scale;
            double offset_Right = (rightPoint - drawingView1.Position.X) / drawingView1.Scale;


            Point2d pt1 = invapp.TransientGeometry.CreatePoint2d(drawingView1.Position.X-100, offset);
            Point2d pt2 = invapp.TransientGeometry.CreatePoint2d(drawingView1.Position.X-100 + (2 * 2.54), offset);
            Point2d pt3 = invapp.TransientGeometry.CreatePoint2d(drawingView1.Position.Y, offset_Bot);
            Point2d pt4 = invapp.TransientGeometry.CreatePoint2d(drawingView1.Position.Y + 2 * 2.54, offset_Bot);

            Point2d pt5 = invapp.TransientGeometry.CreatePoint2d(offsetleft,drawingView1.Position.Y- offset);
            Point2d pt6 = invapp.TransientGeometry.CreatePoint2d(offsetleft, drawingView1.Position.Y - offset + (2 * 2.54));
            Point2d pt7 = invapp.TransientGeometry.CreatePoint2d(offset_Right, drawingView1.Position.Y - offset);
            Point2d pt8 = invapp.TransientGeometry.CreatePoint2d(offset_Right, drawingView1.Position.Y - offset + (2 * 2.54));

            DrawingSketch drawingSketch = drawingView1.Sketches.Add();
            drawingSketch.Edit();
            SketchLine drawingLine = drawingSketch.SketchLines.AddByTwoPoints(pt1,pt2);
            SketchLine drawingLine2 = drawingSketch.SketchLines.AddByTwoPoints(pt3, pt4);

            SketchLine drawingLine3 = drawingSketch.SketchLines.AddByTwoPoints(pt5, pt6);
            SketchLine drawingLine4 = drawingSketch.SketchLines.AddByTwoPoints(pt7, pt8);
            drawingSketch.ExitEdit();

            GeometryIntent geoIntent1 = oSheet.CreateGeometryIntent(drawingLine, PointIntentEnum.kCenterPointIntent);
            GeometryIntent geoIntent2 = oSheet.CreateGeometryIntent(drawingLine2, PointIntentEnum.kCenterPointIntent);
            GeometryIntent geoIntent3 = oSheet.CreateGeometryIntent(drawingLine3, PointIntentEnum.kCenterPointIntent);
            GeometryIntent geoIntent4 = oSheet.CreateGeometryIntent(drawingLine4, PointIntentEnum.kCenterPointIntent);

            Point2d ptv = invapp.TransientGeometry.CreatePoint2d(drawingView1.Position.X - 10 + 2 * 2.54, drawingView1.Position.Y);

            Point2d pth = invapp.TransientGeometry.CreatePoint2d(drawingView1.Position.X, drawingView1.Position.Y-6);

            GeneralDimensions dims = oSheet.DrawingDimensions.GeneralDimensions;
            GeneralDimension dim = (GeneralDimension)dims.AddLinear(ptv, geoIntent1, geoIntent2, DimensionTypeEnum.kVerticalDimensionType);

            GeneralDimension dim1 = (GeneralDimension)dims.AddLinear(pth, geoIntent3, geoIntent4, DimensionTypeEnum.kHorizontalDimensionType);

            // DRW RENAME & SAVEAS
            string newFileName = "D:\\onedrive shyam\\LTask\\output.idw";
            try
            {
                drawingDoc.SaveAs(newFileName, true);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
            // ALL DOC CLOSE
            drawingDoc.Close(true);
            partDoc.Close();


            //invapp.Quit();  
            //System.Windows.Forms.Application.Exit();
            this.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            btn_ClickMe_Click(sender, e);
        }
    }

}
