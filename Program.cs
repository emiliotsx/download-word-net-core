using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordGeneratorAPI
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string filePath = "doc.docx"; // Ruta al archivo existente

            var dictionary = new Dictionary<string, string>
            {
                { "Describe the methods used to gather information, including interviews, document reviews, and data analysis techniques.", "The methods used to gather information in the investigation included:\n1. **Interviews**: Detailed interviews were conducted with relevant personnel, including employees who were directly involved in the procurement processes. Notable interviewees included Joe Smith (Procurement Manager) and Becky Hall (Procurement Head). Interview summaries provided significant statements and insights from these interviews, such as explanations of procurement processes and the rationale behind vendor selections.\n2. **Document Reviews**: A comprehensive review of procurement and financial records was undertaken. This included examining relevant documents such as purchase orders, invoices, bidding records, and selection documents. Specific attention was given to the approval processes and compliance with company policies.\n3. **Data Analysis Techniques**: The investigation involved analyzing procurement and payment-related documents to identify any indications of potentially problematic activity or deviations from established policies and procedures. This included looking for patterns in bidding processes and examining financial records for irregularities.\n4. **Targeted Searches**: The investigation included targeted searches of email and chat correspondence related to the individuals involved in the allegations, focusing on communications that may indicate improper conduct or conflicts of interest.\n5. **Public Searches**: Public searches were conducted to identify the ownership and background of the vendors involved, specifically ABC Corp and XYZ Inc.\nThese methods are outlined in the **Methodology** section of the investigation report (III), along with specific subsections detailing each approach." },
                { "Explain the broader context in which the investigation is situated, including industry trends and organizational challenges.", "This information cannot be found on the supporting documentation." },
                {"Provide a brief overview of the investigation's purpose, key findings, and recommendations. This section should summarize the main points of the report for quick reference","The investigation was initiated following an anonymous whistleblower report alleging improper conduct involving Joe Smith (Procurement) and Karen Jones (Supply Chain) at the Chicago, IL facility. The allegations centered on a potential vendor kickback scheme involving ABC Corp and XYZ Inc., which included claims of preferential treatment in procurement processes and the approval of invoices for non-rendered services. This context prompted a comprehensive investigation to assess the validity of these claims and the adherence to company policies and procedures (Investigation Report, Case ID: 12345, Section I - Background and Objectives)."},
                {"Discuss the context that led to the investigation, including any relevant events or triggers.","Overview of the Investigation:\nPurpose:\nThe investigation was initiated following an anonymous whistleblower report alleging improper conduct involving Joe Smith (Procurement) and Karen Jones (Supply Chain) at the Chicago, IL facility. The allegations centered on a potential vendor kickback scheme with ABC Corp and XYZ Inc., which involved preferential treatment in procurement processes and the approval of invoices for non-rendered services.\nKey Findings:\n1. Documentary Evidence:\n- An examination of Invoice 1592548 from ABC Corp revealed a $50,000 charge for 'consulting advice,' a service not typically provided by the vendor. This invoice was approved by Karen Jones without adhering to the required approval process for amounts exceeding $40,000.\n- The selection documents for XYZ Inc. showed insufficient business justification for selecting the vendor over lower-priced bids. The justification labeled as 'best choice' lacked detailed explanation.\n2. Interview Summaries:\n- Sam Walsh acknowledged unawareness of the specifics behind the 'consulting advice' service and recognized that the invoice approval did not follow established company procedures.\n- Joe Smith admitted to a personal relationship with XYZ Inc., owned by his wife, and failed to disclose this conflict of interest.\n3. Compliance with Policies and Procedures:\n- Significant deviations from the Corporate Direct Procurement Policy and the Corporate Code of Conduct Policy were identified, particularly regarding Joe Smith's undisclosed conflict of interest and improper invoice approvals.\nRecommendations:\n1. Initiate a comprehensive financial audit focusing on all transactions and contracts involving ABC Corp and XYZ Inc. to uncover any further irregularities.\n2. Revise and reinforce the procurement and invoice approval processes, ensuring that all employees are aware of and adhere to the proper procedures.\n3. Implement targeted training sessions on ethics, conflict of interest, and procurement policies for all employees, particularly those involved in procurement and financial roles.\n4. Consider appropriate disciplinary actions against Joe Smith and Karen Jones, up to and including termination, for their violations of company policies.\nThis summary encapsulates the main points of the investigation report, which is essential for quick reference."},
                {"Provide a detailed description of external, employee, and operational fraud risks within corporations.","This information cannot be found on the supporting documentation."},
                {"Outline the specific objectives that the investigation aimed to achieve.","The specific objectives that the investigation aimed to achieve include:\n1. Investigating Allegations: To investigate the allegations made in the anonymous whistleblower report regarding improper conduct involving Joe Smith (Procurement) and Karen Jones (Supply Chain) at the Chicago, IL facility, specifically focusing on a potential vendor kickback scheme with ABC Corp and XYZ Inc.\n2. Assessment of Procurement Processes: To assess the procurement processes for any preferential treatment granted to ABC Corp and XYZ Inc., particularly in the approval of invoices for non-rendered services.\n3. Compliance Evaluation: To evaluate the adherence to company policies and procedures, including the Corporate Direct Procurement Policy and the Corporate Code of Conduct Policy, particularly in relation to conflicts of interest and invoice approvals.\n4. Data Gathering: To gather relevant documentation, including purchase orders, invoices, payments, bidding records, and contracts involving ABC Corp and XYZ Inc., and to conduct interviews with involved personnel.\n5. Identifying Irregularities: To identify any irregularities or deviations from established processes and policies that may indicate unethical or illegal conduct.\n6. Recommendations for Future Prevention: To formulate recommendations for preventing future occurrences of similar issues, including potential disciplinary actions against involved employees and revisions to current policies and training programs.\nThese objectives are outlined in the 'Background and Objectives' section of the investigation report."}
            };

            ModifyWordDocument(filePath, dictionary);
        }

        private static void ModifyWordDocument(string filePath, Dictionary<string, string> dictionary)
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;


                    List<string> consolidatedBracketTexts = new List<string>();
                    List<Paragraph> paragraphsToRemove = new List<Paragraph>();
                    StringBuilder currentBracketText = new StringBuilder();
                    bool isInsideBrackets = false;
                    int quantityParagraphsJumpLine = 0; // Se utiliza para determinar cantos saltos de linea tiene el texto que esta dentro de corchetes, de esta manera sabremos cuantos parrafos remover

                    var paragraphs = body.Elements<Paragraph>().ToList();
                    for (int i = 0; i < paragraphs.Count; i++)
                    {
                        string paragraphText = string.Join(" ", paragraphs[i].Descendants<Text>().Select(t => t.Text));

                        if (paragraphText.StartsWith("["))
                        {
                            quantityParagraphsJumpLine = 1;
                        }
                        else if (quantityParagraphsJumpLine >= 1)
                        {
                            quantityParagraphsJumpLine++;
                        }

                        foreach (char c in paragraphText)
                        {
                            if (c == '[')
                            {
                                isInsideBrackets = true;
                                currentBracketText.Clear(); // Reiniciar texto actual
                            }

                            if (isInsideBrackets)
                            {
                                currentBracketText.Append(c);
                            }

                            if (c == ']')
                            {
                                isInsideBrackets = false;
                                string consolidatedText = CleanText(currentBracketText.ToString());

                                foreach (var entry in dictionary)
                                {
                                    string searchText = CleanText(entry.Key);
                                    string insertText = entry.Value.Replace("*", ""); // Texto a insertar
                                    if (string.Equals(consolidatedText, searchText, StringComparison.InvariantCulture))
                                    {
                                        Console.WriteLine($"pasting answer for: {searchText}");
                                        // guarda los parrafos que deben de removerse
                                        for (int j = 0; j < quantityParagraphsJumpLine; j++)
                                        {
                                            // esto es para que no borre la respuesta, ya que al pegar la respuesta lo toma como un mismo parrafo (parrafo que coincide + respuesta)
                                            if (j > 0)
                                            {
                                                int indexToRemove = i - j;
                                                paragraphsToRemove.Add(paragraphs[indexToRemove]);
                                            }
                                        }

                                        // esto es para que siempre borre el parrafo que coincide y no borre la respuesta
                                        paragraphs[i].RemoveAllChildren<Run>();

                                        AddBreakToParagraph(paragraphs[i], 1);
                                        AddTextToParagraph(paragraphs[i], insertText);
                                        break;
                                    }
                                }
                                quantityParagraphsJumpLine = 0;
                                currentBracketText.Clear();

                            }
                        }
                    }

                    // despues de pegar la informacion, borra los parrafos de pregunta, aun asi tengan saltos de linea los borra todos
                    foreach (var paragraph in paragraphsToRemove.Distinct())
                    {
                        paragraph.Remove();
                    }


                    wordDoc.MainDocumentPart.Document.Save();
                }

                Console.WriteLine("Document modified successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error modifying the document: {ex.Message}");
            }
        }

        private static string CleanText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            // Eliminar saltos de línea y espacios redundantes
            text = text.Replace("\n", " ").Replace("\r", " ").Trim();

            // Reemplazar múltiples espacios por uno solo
            text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");

            // Eliminar símbolos no deseados (como números decimales y otros)
            text = System.Text.RegularExpressions.Regex.Replace(text, @"[^\w\s\[\]]", "");

            return new string(text
                .Where(c => !char.IsPunctuation(c) && c != '\n' && c != '\r') // Eliminar puntuación
                .Select(c => char.ToLower(c)) // Convertir a minúsculas
                .ToArray())
                .Trim();
        }

        private static void AddBreakToParagraph(Paragraph paragraph, int count)
        {
            for (int i = 0; i < count; i++)
            {
                paragraph.Append(new Run(new Break()));
            }
        }

        private static void AddTextToParagraph(Paragraph paragraph, string content)
        {
            foreach (var line in content.Split('\n'))
            {
                Run run = new Run();
                run.Append(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
                paragraph.Append(run);
                paragraph.Append(new Run(new Break())); // Salto de línea
            }
        }

    }
}
