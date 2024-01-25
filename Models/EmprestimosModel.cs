using System.ComponentModel.DataAnnotations;

namespace EmprestimoLivros.Models
{
    public class EmprestimosModel
    {
        [Key]
        public int Id { get; set; }
        [Required(ErrorMessage = "Digite o nome do recebedor")]
        public string Recebedor { get; set; }
        [Required(ErrorMessage = "Digite o nome do fornecedor")]
        public string Fornecedor { get; set; }
        [Required(ErrorMessage = "Digite o nome do livro")]
        public string LivroEmprestado { get; set; }
        public DateTime DataUltimaAtualizacao { get; set; } = DateTime.Now;
    }
}
