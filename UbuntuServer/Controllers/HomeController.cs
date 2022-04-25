using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using System.Linq;
using Microsoft.AspNetCore.Http;
using TemplateEngine.Docx;


namespace UbuntuServer.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Produces("application/json")]
    public class HomeController : Controller
    {
        [HttpPost("import")]
        public IActionResult Import(IFormFile file)
        {
            return CreatedAtAction(nameof(Import), ImportFile(file));
        }

        private List<string> ImportFile(IFormFile file)
        {
            using var stream = new System.IO.MemoryStream();
            file.CopyTo(stream);
            stream.Position = 0;

            using var workBook = new XLWorkbook(stream);
            var workSheet = workBook.Worksheets.FirstOrDefault();

            IEnumerable<IXLRow> usedRows = workSheet.Rows(2, workSheet.RowsUsed().Count());            
            List<string> collection = new List<string>();
            collection.Add(usedRows.FirstOrDefault().Cell(1).CachedValue.ToString());
            collection.Add(usedRows.FirstOrDefault().Cell(2).CachedValue.ToString());
            collection.Add(usedRows.FirstOrDefault().Cell(3).CachedValue.ToString());

            return collection;

        }

        [HttpPost("CreateDocx")]
        public async Task<IActionResult> CreateDocxFromTemplate(IFormFile file) 
        {
            CreateDocx(file);
            return Ok();
        }

        private void CreateDocx(IFormFile file)
        {				
            System.IO.File.Delete("OutputDocument.docx");
            System.IO.File.Copy("template.docx", "output.docx");
		
	        var valuesToFill = new Content(
		        new FieldContent("Report date", DateTime.Now.ToString()));

		    using(var outputDocument = new TemplateProcessor("output.docx")
				.SetRemoveContentControls(true))
            {
                outputDocument.FillContent(valuesToFill);
                outputDocument.SaveChanges();
            } 
		}
	
    }
}