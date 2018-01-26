using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebServiceVisio.AccDatos
{
    public class AccesoDatos
    {
        private static NpgsqlConnection conn = null;
        private string connstring = "Server=localhost;Port=5432;User Id=postgres;Password=1234;Database=visio;";
        public AccesoDatos()
        {
            try
            {
                conn = new NpgsqlConnection(connstring);
                conn.Open();
            }
            catch (Exception e)
            {
                throw;   
            }
        }
        public Boolean InsertFigura(string nombre_figura,double cor_x,double cor_y,string texto)
        {
            try
            {
                NpgsqlCommand cmd = new NpgsqlCommand("insert into figura(nombre_figura,cor_x,cor_y,texto) values('" + nombre_figura+ "','" + cor_x.ToString().Replace(',', '.') + "','" + cor_y.ToString().Replace(',', '.') + "','" + texto+"')", conn);
                cmd.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }
    }
}