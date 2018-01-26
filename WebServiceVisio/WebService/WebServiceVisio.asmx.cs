using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using WebServiceVisio.Visio;

namespace WebServiceVisio.WebService
{
    /// <summary>
    /// Descripción breve de WebServiceVisio
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente. 
    // [System.Web.Script.Services.ScriptService]
    public class WebServiceVisio : System.Web.Services.WebService
    {
        public VisioShapes shape=null;
        public WebServiceVisio()
        {
            shape = new VisioShapes();
        }
        [WebMethod]
        public string GenerarRectangulo(double cor_x,double cor_y,string texto)
        {
            return shape.GenerateRectangle(cor_x,cor_y,texto);
        }
        [WebMethod]
        public string GenerarEstrella(double cor_x, double cor_y, string texto)
        {
            return shape.GenerateStar(cor_x, cor_y, texto);
        }
        [WebMethod]
        public string GenerarHexagono(double cor_x, double cor_y, string texto)
        {
            return shape.GenerateHexagon(cor_x, cor_y, texto);
        }
        [WebMethod]
        public string GenerarTriangulo(double cor_x, double cor_y, string texto)
        {
            return shape.GenerateTriangle(cor_x, cor_y, texto);
        }
        [WebMethod]
        public string GenerarCirculo(double cor_x, double cor_y, string texto)
        {
            return shape.GenerateCircle(cor_x, cor_y, texto);
        }
        [WebMethod]
        public string GenerarCuadrado3D(double cor_x, double cor_y, string texto)
        {
            return shape.Generate3DBox(cor_x, cor_y, texto);
        }
    }
}
