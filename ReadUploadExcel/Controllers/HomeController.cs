using ReadUploadExcel.Models;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ReadUploadExcel.Controllers
{
    public class HomeController : Controller
    {
        private Context db = new Context();

        // GET: Home
        public ActionResult Index()
        {
            ViewBag.lista = db.Pessoas.ToList();

            return View();
        }

        public ActionResult LoadExcel()
        {
            return View();
        }

        [HttpPost]
        public ActionResult LoadExcel(HttpPostedFileBase fileExcel)
        {
            try
            {
                if (fileExcel != null)
                {
                    string filename = fileExcel.FileName;
                    string caminho = System.IO.Path.Combine(Server.MapPath("~/Storage/"), filename);

                    fileExcel.SaveAs(caminho);
                    List<Pessoa> excel = GetDataFromExcel(caminho);

                    if (excel != null)
                    {
                        foreach (var item in excel)
                        {
                            var pessoa = new Pessoa
                            {
                                Nome = item.Nome,
                                Sobrenome = item.Sobrenome,
                                Email = item.Email
                            };

                            db.Pessoas.Add(pessoa);
                            db.SaveChanges();
                        }

                        ViewBag.Sucesso = "Inserido com sucesso!";
                    }
                }
                else
                {
                    ViewBag.Erro = "Deu ruim";
                }

                return RedirectToAction("Index");
            }
            catch (Exception ex)
            {
                ViewBag.Erro = ex.Message;
                return View();
            }
        }

        public List<Pessoa> GetDataFromExcel(string filepath)
        {
            string connectionstring = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                   "Data Source=" + filepath + ";" + "Extended Properties=" + "\"" + "Excel 12.0;HDR=YES;" + "\"";

            OleDbConnection connection = new OleDbConnection(connectionstring);
            try
            {
                List<Pessoa> pessoas = new List<Pessoa>();

                string sql = "select *from [Plan1$]";

                OleDbCommand command = new OleDbCommand(sql, connection);

                connection.Open();

                OleDbDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    pessoas.Add(new Pessoa
                    {
                        Id = Convert.ToInt32(reader["Id"]),
                        Nome = reader["Nome"].ToString(),
                        Sobrenome = reader["Sobrenome"].ToString(),
                        Email = reader["Email"].ToString()
                    });
                }

                if (pessoas.Count > 0) return pessoas;
                else return null;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                connection.Close();
            }
        }
    }
}