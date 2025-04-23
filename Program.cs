using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordGeneratorAPI
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Generar documento Word automáticamente al iniciar el proyecto
            GenerateWordDocument();

            // Configuración del host para ejecutar el servidor
            var builder = WebApplication.CreateBuilder(args);
            var app = builder.Build();

            app.MapGet("/", () => "Word generation project running...");
            app.Run();
        }

        private static void GenerateWordDocument()
        {
            try
            {
                // Nombre del archivo generado
                string fileName = $"GeneratedDocument_{DateTime.Now:yyyyMMddHHmmss}.docx";
                string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

                // Diccionario proporcionado
                var data = new Dictionary<string, string>
                {
                    { "I. Background and Objectives\n- Description of the initial allegations and the objectives of the investigation.\n\n", "I. Background and Objectives\n\nThe investigation was initiated following an anonymous whistleblower report alleging improper conduct involving Joe Smith (Procurement) and Karen Jones (Supply Chain) at the Chicago, IL facility. Specifically, the allegations centered on a potential vendor kickback scheme involving ABC Corp and XYZ Inc., which included preferential treatment in procurement processes and the approval of invoices for non-rendered services. The objective of the investigation is to thoroughly examine these allegations, review relevant procurement and financial records, assess compliance with company policies, and identify any potential conflicts of interest or unethical practices within the procurement process." },
                    { "II. Investigation Team\n- List of team members and their roles in the investigation.\n\n", "II. Investigation Team\n- Christine Valentino, Investigation Lead\n- Joe Smith, Procurement\n- Karen Jones, Supply Chain\n- Becky Hall, Procurement Head\n- Mark Linder, Department Head\n\n(Note: The specific roles of Joe Smith and Karen Jones in the investigation are implied, as they are mentioned in relation to the allegations, but their exact titles in the investigation context are not explicitly stated.)" },
                    { "III. Methodology\n- Summary of the investigation workplan.\n\n", "### III. Methodology  \n**Summary of the Investigation Workplan**  \nThe investigation workplan for Case ID: 12345, led by Christine Valentino, includes the following steps:\n\n1. **Contact Whistleblower**: \n   - Reiterate the company’s commitment to investigate the allegations.\n   - Emphasize the company’s anti-retaliation policy.\n   - Inquire if the whistleblower is willing to speak with the investigator via phone or email.\n\n2. **Gather Information on Alleged Vendors and Employees**: \n   - Collect the following:\n     - List of purchase orders, invoices, and payments for ABC Corp and XYZ Inc.\n     - Bidding records involving ABC Corp and XYZ Inc.\n     - Contracts with ABC Corp and XYZ Inc.\n     - Organizational charts for procurement and supply chain.\n   - Investigate historical matters to determine if there have been past allegations against Joe Smith, Karen Jones, ABC Corp, or XYZ Inc.\n\n3. **Place Legal Hold**: \n   - Implement a legal hold on Joe Smith and Karen Jones’ server data, including emails and chats.\n\n4. **Review Relevant Policies and Procedures**: \n   - Review policies such as the Code of Conduct, Conflicts of Interest, and Procurement Policy.\n\n5. **Determine Disclosed Conflicts of Interest**: \n   - Assess whether any conflicts of interest were disclosed by the involved parties.\n\n6. **Conduct Targeted Email and Chat Searches**: \n   - Examine correspondence between Joe Smith, Karen Jones, and both ABC Corp and XYZ Inc. for evidence of alleged misconduct.\n\n7. **Public Searches**: \n   - Identify the owners of ABC Corp and XYZ Inc. through public information.\n\n8. **Consult HR**: \n   - Discuss the information collected to date and decide whether HR should participate in interviews.\n\n9. **Conduct Interviews**: \n   - Interview relevant individuals, including those from Finance, department heads, and the alleged employees (Joe Smith and Karen Jones).\n\n10. **Draft Investigation Report**: \n   - Prepare the investigation report with findings and recommendations, and discuss it with HR and the remediation committee.\n\nThis workplan is designed to ensure a thorough and confidential investigation in line with the company's policies and procedures." },
                    { "IV. Observations\n1. **Documentary Evidence:**  \n   - Summary of relevant documents reviewed and key information extracted.\n   \n2. **Interview Summaries:**  \n   - Synopses of interviews conducted, highlighting significant statements and insights.\n   \n3. **Compliance with Policies and Procedures:**  \n   - Assessment of adherence to company policies and procedures by the involved parties.\n\n", "1. **Documentary Evidence:**  \\n   - A detailed examination of Invoice 1592548 from ABC Corp revealed a $50,000 charge for \\\"consulting advice,\\\" a service that is not typically provided by the vendor. The invoice was approved by Karen Jones, which bypassed the standard approval process requiring Department Head authorization for amounts exceeding $40,000.  \\n   - Analysis of selection documents for XYZ Inc. indicated insufficient business justification for selecting the vendor over lower-priced bids. The justification labeled as \\\"best choice\\\" lacked a detailed explanation, which raised concerns about the objectivity of the selection process.\\n\\n2. **Interview Summaries:**  \\n   - Sam Walsh expressed unawareness regarding the specifics of the \\\"consulting advice\\\" service and acknowledged that the invoice approval did not adhere to the company’s established procedures.  \\n   - Joe Smith admitted to having a personal relationship with XYZ Inc., owned by his wife, and failed to disclose this conflict of interest. He could not provide a valid reason for instructing XYZ Inc. to increase an invoice by $5,000.  \\n   - Becky Hall recognized that the business justification for selecting XYZ Inc. was insufficient and committed to more diligently reviewing selection documents in the future.  \\n   - Mark Linder was surprised by the investigation’s findings, stating he had no prior knowledge of personal ties between company employees and the vendors or any issues with the vendors' performance.\\n\\n3. **Compliance with Policies and Procedures:**  \\n   - The investigation identified significant deviations from the Corporate Direct Procurement Policy and the Corporate Code of Conduct Policy. Notably, Joe Smith's failure to disclose his conflict of interest with XYZ Inc. and the improper approval of invoices highlighted gaps in adherence to and enforcement of these policies.\",\r\n    \" V. Conclusions\\n- Statement on whether the allegations could be or could not be substantiated.\\n\\n\": \"The investigation substantiated parts of the whistleblower's allegations, particularly concerning Joe Smith's undisclosed conflict of interest with XYZ Inc. and the irregular approval of invoices. Notably, Joe Smith could not provide a valid reason for instructing XYZ Inc. to increase an invoice by $5,000, and Becky Hall recognized the business justification for selecting XYZ Inc. as insufficient. There were significant deviations from the Corporate Direct Procurement Policy and the Corporate Code of Conduct Policy identified during the investigation. These findings suggest a breach of company policies and indicate potential ethical violations in the procurement process. Additionally, some financial records and communications may not have been fully accessible.\",\r\n    \" VI. Recommendations\\n- Specific actions proposed to address the findings, prevent future occurrences, and remediate any harm done.\\n\\n\": \"- Initiate a comprehensive financial audit focusing on all transactions and contracts involving ABC Corp and XYZ Inc. to uncover any further irregularities.\\n- Revise and reinforce the procurement and invoice approval processes, ensuring that all employees are aware of and adhere to the proper procedures.\\n- Implement targeted training sessions on ethics, conflict of interest, and procurement policies for all employees, with a focus on those involved in procurement and financial roles.\\n- Consider appropriate disciplinary actions against Joe Smith and Karen Jones, up to and including termination, for their violations of company policies." },
                    { "VI. Recommendations\n- Specific actions proposed to address the findings, prevent future occurrences, and remediate any harm done.\n\n", "- Initiate a comprehensive financial audit focusing on all transactions and contracts involving ABC Corp and XYZ Inc. to uncover any further irregularities.\n- Revise and reinforce the procurement and invoice approval processes, ensuring that all employees are aware of and adhere to the proper procedures.\n- Implement targeted training sessions on ethics, conflict of interest, and procurement policies for all employees, with a focus on those involved in procurement and financial roles.\n- Consider appropriate disciplinary actions against Joe Smith and Karen Jones, up to and including termination, for their violations of company policies." },
                    { "VII. Limitations\n- Any limitations encountered during the investigation that may affect the findings or conclusions.\n\n", "VII. Limitations  \nThe investigation faced limitations due to the reliance on self-reported information during interviews and the absence of direct evidence from the whistleblower. Additionally, some financial records and communications may not have been fully accessible." }
                };

                // Crear el documento
                CreateWordDocument(filePath, data);
                Console.WriteLine($"Document generated successfully: {filePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating Word document: {ex.Message}");
            }
        }

        private static void CreateWordDocument(string filePath, Dictionary<string, string> data)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                // Crear estilos personalizados
                AddStyles(mainPart);

                // Agregar claves y valores del diccionario
                foreach (var entry in data)
                {
                    string title = entry.Key.Split('\n')[0]; // Tomar solo la primera línea del título
                    AddStyledParagraph(body, title, true, 24);
                    AddStyledParagraph(body, ProcessText(entry.Value), false, 12);
                }

                mainPart.Document.Save();
            }
        }

        private static void AddStyles(MainDocumentPart mainPart)
        {
            StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
            Styles styles = new Styles();

            Style heading1 = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1",
                CustomStyle = true,
                StyleName = new StyleName { Val = "Heading 1" }
            };
            heading1.Append(new BasedOn { Val = "Normal" });
            heading1.Append(new NextParagraphStyle { Val = "Normal" });
            heading1.Append(new StyleRunProperties(new Bold(), new FontSize { Val = "28" }));

            Style heading2 = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading2",
                CustomStyle = true,
                StyleName = new StyleName { Val = "Heading 2" }
            };
            heading2.Append(new BasedOn { Val = "Normal" });
            heading2.Append(new NextParagraphStyle { Val = "Normal" });
            heading2.Append(new StyleRunProperties(new Bold(), new FontSize { Val = "24" }));

            styles.Append(heading1);
            styles.Append(heading2);
            stylePart.Styles = styles;
            stylePart.Styles.Save();
        }

        private static void AddStyledParagraph(Body body, string text, bool isBold, int fontSize)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();

            RunProperties runProperties = new RunProperties();
            if (isBold)
            {
                runProperties.Append(new Bold());
            }
            runProperties.Append(new FontSize { Val = (fontSize * 2).ToString() }); // Font size in half-points
            run.Append(runProperties);

            // Dividir el texto por saltos de línea y procesar cada segmento
            foreach (var line in text.Split(new[] { "\n" }, StringSplitOptions.None))
            {
                run.Append(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
                run.Append(new Break()); // Saltos de línea con el mismo tamaño de fuente
            }

            paragraph.Append(run);
            body.Append(paragraph);
        }

        private static string ProcessText(string text)
        {
            return text.Replace("\\n", "\n");
        }
    }
}
