using MySql.Data.MySqlClient;
using System;
using System.IO;
using System.Linq;
using System.Net.Mail;

namespace ConsoleCobranzas
{
    class Program
    {
        public const string connStr = "datasource=netdte.cl;Database=netdte_informatico;username=netdte_service;password=x~Mv%[l.*w)D;";
        const string _ruta = "//servidorvisoft//AceptaServiceVisoft//visoft//var//ca4xml//output//carpeta-con-pdf";
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("***Sisgen Chile Email Service***");
            Console.WriteLine("....................");

            //SetInsertData("angel.pinilla@sisgenchile.cl", "pruebaangel", "33", "3897", "17558736-5", "filename.exe");

            if (args[4] == "SERVICE")
            {
                try
                {
                    ExecuteOrder66(args[0], args[1], args[2], args[3]);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }

                Console.WriteLine("................");
                Console.WriteLine("Proceso terminado");
            }
            else if (args[4] == "REMINDER")
            {
                //GetReminder();
            }
        }

        private static void ExecuteOrder66(string tipoDte, string folio, string rutCliente, string email)
        {
            string[] _arcPDFFiles = Directory.GetFiles(_ruta, "*.pdf").Select(Path.GetFileName).ToArray();
            bool flag = false;

            foreach (var item in _arcPDFFiles)
            {
                string[] split = item.Split(".");

                if (split[1] == tipoDte && split[2] == folio)
                {
                    Console.WriteLine("item: " + item);
                    string path = Path.Combine(_ruta + "//" + item);
                    Console.WriteLine("................");
                    Console.WriteLine("Sending email.....");
                    flag = true;
                    SendEmail(email, path, tipoDte, folio, rutCliente, "Sisgen Chile - Facturación Electrónica");

                    var data = folio + "- " + tipoDte;

                    SetInsertData(email, data, tipoDte, folio, rutCliente, path);
                    break;
                }
            }
            if (!flag)
                Console.WriteLine("Archivo no encontrado. Ocurrió algún problema al generar su PDF.");

        }

        private static void SendEmail(string email, string file, string tipoDte, string folio, string rutCliente, string subject)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                mail.From = new MailAddress("cobranzas@sisgenchile.cl");
                mail.To.Add(email);
                mail.CC.Add("cobranzas@sisgenchile.cl");
                //mail.CC.Add("angel.pinilla@sisgenchile.cl");
                mail.Subject = subject;
                mail.IsBodyHtml = true;
                mail.Attachments.Add(new Attachment(file));

                string cuerpoHtml = "<p>Estimado cliente de Sisgen Chile, junto con saludar y esperando que tenga un buen día," +
                " enviamos factura n° " + folio + " correspondiente a servicios adquiridos con fecha " + DateTime.Now.ToString("dd MMMM yyyy") +
                " </p>" +
                    "<p style='color:#3396FF'><strong>Datos para transferencia</strong></p>" +
                    "<p><strong>Cuenta Corriente:</strong> 6856578</p>" +
                    "<p><strong>Banco:</strong> Banco Estado</p>" +
                    "<p><strong>Rut:</strong> 76.161.082-1</p>" +
                    "<p><strong>Razón Social:</strong> Sisgen Chile Computación Limitada</p>" +
                    "<p><strong>E-mail:</strong><strong style='color:red'> ENVIAR COMPROBANTE A</strong> cobranzas@sisgenchile.cl</p>" +
                    "<p>Les saluda atentamente,</p>" +
                    "<p><strong>Departamento de Administración y Cobranzas Sisgen Chile Limitada</strong></p>" +
                    "<a href='https://www.sisgenchile.cl'>Visitar Sitio Web Sisgen Chile Limitada</a>" +
                    "<p class='text-muted'>Importante: Si usted tiene regularizada su situación de cobranza con nuestra empresa, omitir este correo.</p>";

                mail.Body = cuerpoHtml;

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("cobranzas@sisgenchile.cl", "ETG7n>5e");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);

                Console.WriteLine("Email enviado");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);

                MySqlConnection conn = new MySqlConnection(connStr);
                MySqlCommand comm = conn.CreateCommand();
                comm.CommandText = string.Format("insert into Log(Descripcion,CorreoCliente,Tipo) " +
                    "values('{0}','{1}','{2}')", ex + " - ", email, "Error Email");
                comm.ExecuteNonQuery();
                conn.Close();
            }
        }

        private static void SetInsertData(string email, string data, string tipodte, string folio, string rutCliente, string fileName)
        {
            Console.WriteLine("................");
            Console.WriteLine("Insert data.....");

            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {

                conn.Open();
                MySqlCommand comm = conn.CreateCommand();

                comm.CommandText = string.Format("insert into MailService(Email,TipoDte, Folio,RutCliente,Archivo) " +
                    "values('{0}','{1}','{2}','{3}','{4}')", email, tipodte, folio, rutCliente, @fileName);

                comm.ExecuteNonQuery();

                comm.CommandText = string.Format("insert into Log(Descripcion,CorreoCliente,Tipo) " +
                 "values('{0}','{1}','{2}')", "Se han guardado nuevos registros en la tabla Mail Service. Datos: " + data + ".", email, "Success");
                comm.ExecuteNonQuery();

                conn.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                MySqlCommand comm = conn.CreateCommand();
                comm.CommandText = string.Format("insert into Log(Descripcion,CorreoCliente,Tipo) " +
                    "values('{0}','{1}','{2}')", ex + " - " + data, email, "Error");
                comm.ExecuteNonQuery();
            }

        }



        private static void GetReminder()
        {
            MySqlConnection conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();

                string sql = "select FechaEnvio,Email,TipoDte,Folio,RutCliente,Archivo from MailService";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                MySqlDataReader rdr = cmd.ExecuteReader();
                while (rdr.Read())
                {
                    DateTime fechaInicial = DateTime.Parse(rdr[0].ToString());

                    var fechaFormateada = fechaInicial.AddDays(4);
                    var today = DateTime.Now.ToString("dd-MM-yyyy");

                    if (fechaFormateada.ToString("dd-MM-yyyy") == today)
                    {
                        SendEmail(rdr[1].ToString(), @rdr[5].ToString(), rdr[2].ToString(),
                          rdr[3].ToString(), rdr[4].ToString(),
                        "¡Aviso vencimiento! Sisgen Chile - Facturación Electrónica");

                        InsertMySqlReader(rdr);
                    }

                }
                rdr.Close();
                conn.Close();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                MySqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = string.Format("insert into Log(Descripcion,CorreoCliente,Tipo) " +
                    "values('{0}','{1}','{2}')", ex, "angel.pinilla@sisgenchile.cl", "Error");
                cmd.ExecuteNonQuery();
            }
            finally
            {
                conn.Close();
            }

        }

        private static void InsertMySqlReader(MySqlDataReader rdr)
        {
            using (MySqlConnection connection = new MySqlConnection(connStr))
            {
                connection.Open();
                string query = string.Format("insert into Log(Descripcion,CorreoCliente,Tipo) " +
               "values('{0}','{1}','{2}')", "Recordatorio enviado a cliente " + rdr[4] + ".",
              rdr[1].ToString(), "Success");

                using (MySqlCommand cmd = new MySqlCommand(query, connection))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}
