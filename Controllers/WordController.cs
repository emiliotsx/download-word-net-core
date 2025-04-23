using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;

namespace WordGeneratorAPI_netcore.Controllers
{
  [ApiController]
  [Route("api/[controller]")]
  public class WordController : ControllerBase
  {
    [HttpPost("generate")]
    public async Task<IActionResult> GenerateWordDocument([FromBody] Dictionary<string, string> data)
    {
      try
      {
        // Nombre temporal del archivo
        string fileName = $"GeneratedDocument_{DateTime.Now:yyyyMMddHHmmss}.docx";
        string tempPath = Path.Combine(Path.GetTempPath(), fileName);

        // Crear el documento
        CreateWordDocument(tempPath, data);

        // Leer el archivo generado y devolverlo como descarga
        var fileBytes = await System.IO.File.ReadAllBytesAsync(tempPath);
        return File(fileBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", fileName);
      }
      catch (Exception ex)
      {
        return BadRequest(new { Message = "Error generating Word document", Details = ex.Message });
      }
    }

    private void CreateWordDocument(string filePath, Dictionary<string, string> data)
    {
      using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
      {
        // Crear el main document part
        MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
        mainPart.Document = new Document();
        Body body = mainPart.Document.AppendChild(new Body());

        // Crear estilos personalizados
        AddStyles(mainPart);

        // Agregar título principal
        AddParagraph(body, "Generated Document", true, "Heading1");

        // Agregar claves y valores del diccionario
        foreach (var entry in data)
        {
          AddParagraph(body, $"{entry.Key}:", true, "Heading2");
          AddParagraph(body, entry.Value);
        }

        mainPart.Document.Save();
      }
    }

    private void AddStyles(MainDocumentPart mainPart)
    {
      StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
      Styles styles = new Styles();

      // Crear un estilo Heading1
      Style heading1 = new Style()
      {
        Type = StyleValues.Paragraph,
        StyleId = "Heading1",
        CustomStyle = true,
        StyleName = new StyleName() { Val = "Heading 1" }
      };
      heading1.Append(new BasedOn() { Val = "Normal" });
      heading1.Append(new NextParagraphStyle() { Val = "Normal" });
      heading1.Append(new StyleRunProperties(new Bold(), new FontSize() { Val = "28" }));

      // Crear un estilo Heading2
      Style heading2 = new Style()
      {
        Type = StyleValues.Paragraph,
        StyleId = "Heading2",
        CustomStyle = true,
        StyleName = new StyleName() { Val = "Heading 2" }
      };
      heading2.Append(new BasedOn() { Val = "Normal" });
      heading2.Append(new NextParagraphStyle() { Val = "Normal" });
      heading2.Append(new StyleRunProperties(new Bold(), new FontSize() { Val = "24" }));

      styles.Append(heading1);
      styles.Append(heading2);
      stylePart.Styles = styles;
      stylePart.Styles.Save();
    }

    private void AddParagraph(Body body, string text, bool isBold = false, string styleId = null)
    {
      Paragraph paragraph = new Paragraph();
      Run run = new Run();

      // Aplicar estilos
      RunProperties runProperties = new RunProperties();
      if (isBold) runProperties.Append(new Bold());
      run.Append(runProperties);

      run.Append(new Text(text));
      paragraph.Append(run);

      // Agregar estilo de párrafo
      if (styleId != null)
      {
        ParagraphProperties paragraphProperties = new ParagraphProperties();
        paragraphProperties.ParagraphStyleId = new ParagraphStyleId() { Val = styleId };
        paragraph.PrependChild(paragraphProperties);
      }

      body.Append(paragraph);
    }
  }

}
