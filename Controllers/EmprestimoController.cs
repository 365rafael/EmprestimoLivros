using ClosedXML.Excel;
using EmprestimoLivros.Data;
using EmprestimoLivros.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Data;

namespace EmprestimoLivros.Controllers
{
    public class EmprestimoController : Controller
    {

        private readonly ApplicationDbContext _db;

        public EmprestimoController(ApplicationDbContext context)
        {
            _db = context;
        }

        public IActionResult Index()
        {
            IEnumerable<EmprestimosModel> emprestimos = _db.Emprestimos;

            return View(emprestimos);
        }

        public IActionResult Cadastrar()
        {
            return View();
        }

        [HttpGet]
        public IActionResult Editar(int? id)
        {
            if (id == null || id == 0)
            {
                return NotFound();
            }

            EmprestimosModel emprestimo = _db.Emprestimos.FirstOrDefault(x => x.Id == id);

            if (emprestimo == null)
            {
                return NotFound();
            }

            return View(emprestimo);
        }

        [HttpGet]
        public IActionResult Excluir(int id)
        {
            if (id == null || id == 0)
            {
                return NotFound();
            }

            EmprestimosModel emprestimo = _db.Emprestimos.FirstOrDefault(x => x.Id == id);

            if (emprestimo == null)
            {
                return NotFound();
            }

            return View(emprestimo);
        }

        public IActionResult Exportar()
        {
            var dados = GetDados();

            //using é fechado assim que o código passa por ele, não fica executando
            using (XLWorkbook workbook = new XLWorkbook())
            {
                workbook.AddWorksheet(dados, "Dados Empréstimos");

                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.SaveAs(ms);
                    return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spredsheet.sheet", "Emprestimo.xls");
                }
            }

        }

        private DataTable GetDados()
        {
            DataTable dataTable = new DataTable();

            dataTable.TableName = "Dados empréstimos";
            dataTable.Columns.Add("Recebedor", typeof(string));
            dataTable.Columns.Add("Fornecedor", typeof(string));
            dataTable.Columns.Add("Livro", typeof(string));
            dataTable.Columns.Add("Data Última Atualizacao", typeof(DateTime));

            var dados = _db.Emprestimos.ToList();

            if (dados.Count > 0)
            {
                dados.ForEach(emprestimo =>
                {
                    dataTable.Rows.Add(
                        emprestimo.Recebedor,
                        emprestimo.Fornecedor,
                        emprestimo.LivroEmprestado,
                        emprestimo.DataUltimaAtualizacao
                        );
                });
            }

            return dataTable;
        }

        [HttpPost]
        public IActionResult Importar()
        {
            if (ModelState.IsValid)
            {
                try
                {
                    var arquivo = Request.Form.Files[0];

                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    if (arquivo != null && arquivo.Length > 0)
                    {
                        using (var stream = new MemoryStream())
                        {
                            arquivo.CopyTo(stream);
                            using (var package = new ExcelPackage(stream))
                            {
                                var worksheet = package.Workbook.Worksheets[0];

                                int totalRows = worksheet.Dimension.Rows;

                                for (int i = 2; i <= totalRows; i++)
                                {
                                    EmprestimosModel emprestimos = new EmprestimosModel
                                    {
                                        Recebedor = worksheet.Cells[i, 1].Value.ToString(),
                                        Fornecedor = worksheet.Cells[i, 2].Value.ToString(),
                                        LivroEmprestado = worksheet.Cells[i, 3].Value.ToString(),
                                        DataUltimaAtualizacao = DateTime.Parse(worksheet.Cells[i, 4].Value.ToString()),
                                    };

                                    _db.Emprestimos.Add(emprestimos);
                                }

                                _db.SaveChanges();
                            }
                        }
                        return RedirectToAction("Index");
                    }

                    return View();
                }
                catch (Exception ex)
                {

                    throw new Exception(ex.Message);
                }
            }
            else
            {
                return RedirectToAction("Index");
            }
        }

        [HttpPost]
        public IActionResult Cadastrar(EmprestimosModel emprestimos)
        {
            if (ModelState.IsValid)
            {
                _db.Emprestimos.Add(emprestimos);
                _db.SaveChanges();

                TempData["MensagemSucesso"] = "Cadastro realizado com sucesso!";

                return RedirectToAction("Index");
            }

            return View();
        }

        [HttpPost]
        public IActionResult Editar(EmprestimosModel emprestimo)
        {
            if (ModelState.IsValid)
            {
                _db.Emprestimos.Update(emprestimo);
                _db.SaveChanges();

                TempData["MensagemSucesso"] = "Edição realizada com sucesso!";

                return RedirectToAction("Index");
            }

            TempData["MensagemErro"] = "Algo deu errado";


            return View(emprestimo);
        }

        [HttpPost]
        public IActionResult Excluir(EmprestimosModel emprestimos)
        {
            if (emprestimos == null)
            {
                return NotFound();
            }

            _db.Emprestimos.Remove(emprestimos);
            _db.SaveChanges();

            TempData["MensagemSucesso"] = "Remoção realizada com sucesso!";

            return RedirectToAction("Index");
        }
    }
}
