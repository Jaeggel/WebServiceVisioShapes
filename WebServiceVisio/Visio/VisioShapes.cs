using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using WebServiceVisio.AccDatos;

namespace WebServiceVisio.Visio
{
    /// <summary>
    /// Clase para la generación de las figuras en Visio 2010
    /// </summary>
    public class VisioShapes
    {
        Microsoft.Office.Interop.Visio.Document visioStencil = null;
        Microsoft.Office.Interop.Visio.Page visioPage = null;
        AccesoDatos db = null;
        public VisioShapes()
        {
            db = new AccesoDatos();
            Microsoft.Office.Interop.Visio.Application Application = new Microsoft.Office.Interop.Visio.Application();
            Application.Documents.Add("");
            Microsoft.Office.Interop.Visio.Documents visioDocs = Application.Documents;
            visioStencil = visioDocs.OpenEx("Basic Shapes.vss",
                (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenDocked);
            visioPage = Application.ActivePage;
        }
        public string GenerateRectangle(double x, double y, string mensaje)
        {
            try
            {
                Microsoft.Office.Interop.Visio.Master visioRectMaster = visioStencil.Masters.get_ItemU(@"Rectangle");
                Microsoft.Office.Interop.Visio.Shape visioRectShape = visioPage.Drop(visioRectMaster, x, y);
                visioRectShape.Text = @mensaje + ".";
                db.InsertFigura("Rectángulo",x,y,mensaje);
            }
            catch (Exception e)
            {
                return "¡ERROR!: " + e.Message;
            }
            return "¡Generación Correcta!";
        }
        public string GenerateStar(double x, double y, string mensaje)
        {
            try
            {
                Microsoft.Office.Interop.Visio.Master visioStarMaster = visioStencil.Masters.get_ItemU(@"Star 7");
                Microsoft.Office.Interop.Visio.Shape visioStarShape = visioPage.Drop(visioStarMaster, x, y);
                visioStarShape.Text = @mensaje + ".";
                db.InsertFigura("Estrella 7 Puntas", x, y, mensaje);
            }
            catch (Exception e)
            {
                return "¡ERROR!: " + e.Message;
            }
            return "¡Generación Correcta!";
        }
        public string GenerateHexagon(double x, double y, string mensaje)
        {
            try
            {
                Microsoft.Office.Interop.Visio.Master visioHexagonMaster = visioStencil.Masters.get_ItemU(@"Hexagon");
                Microsoft.Office.Interop.Visio.Shape visioHexagonShape = visioPage.Drop(visioHexagonMaster, x, y);
                visioHexagonShape.Text = @mensaje + ".";
                db.InsertFigura("Hexágono", x, y, mensaje);
            }
            catch (Exception e)
            {
                return "¡ERROR!: " + e.Message;
            }
            return "¡Generación Correcta!";
        }
        public string GenerateCircle(double x, double y, string mensaje)
        {
            try
            {
                Microsoft.Office.Interop.Visio.Master visioCircleMaster = visioStencil.Masters.get_ItemU(@"Circle");
                Microsoft.Office.Interop.Visio.Shape visioCircleShape = visioPage.Drop(visioCircleMaster, x, y);
                visioCircleShape.Text = @mensaje + ".";
                db.InsertFigura("Círculo", x, y, mensaje);
            }
            catch (Exception e)
            {
                return "¡ERROR!: " + e.Message;
            }
            return "¡Generación Correcta!";
        }
        public string GenerateTriangle(double x, double y, string mensaje)
        {
            try
            {
                Microsoft.Office.Interop.Visio.Master visioTriangleMaster = visioStencil.Masters.get_ItemU(@"Triangle");
                Microsoft.Office.Interop.Visio.Shape visioTriangleShape = visioPage.Drop(visioTriangleMaster, x, y);
                visioTriangleShape.Text = @mensaje + ".";
                db.InsertFigura("Triángulo", x, y, mensaje);
            }
            catch (Exception e)
            {
                return "¡ERROR!: " + e.Message;
            }
            return "¡Generación Correcta!";
        }
        public string Generate3DBox(double x, double y, string mensaje)
        {
            try
            {
                Microsoft.Office.Interop.Visio.Master visioOtherMaster = visioStencil.Masters.get_ItemU(@"3-D box");
                Microsoft.Office.Interop.Visio.Shape visioOtherShape = visioPage.Drop(visioOtherMaster, x, y);
                visioOtherShape.Text = @mensaje + ".";
                db.InsertFigura("Cuadrado 3D", x, y, mensaje);
            }
            catch (Exception e)
            {
                return "¡ERROR!: " + e.Message;
            }
            return "¡Generación Correcta!";
        }
    }
}