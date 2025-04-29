using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
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
                    { "I. Background and Objectives\n- Description of the initial allegations and the objectives of the investigation.\n\n", "I. Background and Objectives\n\nThe investigation was initiated following an anonymous whistleblower report alleging improper conduct involving Joe Smith (Procurement) and Karen Jones (Supply Chain) at the Chicago, IL facility. Specifically, the allegations centered on a potential vendor kickback scheme involving ABC Corp and XYZ Inc., which included preferential treatment in procurement processes and the approval of invoices for non-rendered services. The objective of the investigation is to thoroughly examine these allegations, review relevant procurement and financial records, assess compliance with company policies, and identify any potential conflicts of interest or unethical practices within the procurement process." },
                    { "II. Investigation Team\n- List of team members and their roles in the investigation.\n\n", "II. Investigation Team\n- Christine Valentino, Investigation Lead\n- Joe Smith, Procurement\n- Karen Jones, Supply Chain\n- Becky Hall, Procurement Head\n- Mark Linder, Department Head\n\n(Note: The specific roles of Joe Smith and Karen Jones in the investigation are implied, as they are mentioned in relation to the allegations, but their exact titles in the investigation context are not explicitly stated.)" },
                    { "III. Methodology\n- Summary of the investigation workplan.\n\n", "### III. Methodology  \n**Summary of the Investigation Workplan**  \nThe investigation workplan for Case ID: 12345, led by Christine Valentino, includes the following steps:\n\n1. **Contact Whistleblower**: \n   - Reiterate the company’s commitment to investigate the allegations.\n   - Emphasize the company’s anti-retaliation policy.\n   - Inquire if the whistleblower is willing to speak with the investigator via phone or email.\n\n2. **Gather Information on Alleged Vendors and Employees**: \n   - Collect the following:\n     - List of purchase orders, invoices, and payments for ABC Corp and XYZ Inc.\n     - Bidding records involving ABC Corp and XYZ Inc.\n     - Contracts with ABC Corp and XYZ Inc.\n     - Organizational charts for procurement and supply chain.\n   - Investigate historical matters to determine if there have been past allegations against Joe Smith, Karen Jones, ABC Corp, or XYZ Inc.\n\n3. **Place Legal Hold**: \n   - Implement a legal hold on Joe Smith and Karen Jones’ server data, including emails and chats.\n\n4. **Review Relevant Policies and Procedures**: \n   - Review policies such as the Code of Conduct, Conflicts of Interest, and Procurement Policy.\n\n5. **Determine Disclosed Conflicts of Interest**: \n   - Assess whether any conflicts of interest were disclosed by the involved parties.\n\n6. **Conduct Targeted Email and Chat Searches**: \n   - Examine correspondence between Joe Smith, Karen Jones, and both ABC Corp and XYZ Inc. for evidence of alleged misconduct.\n\n7. **Public Searches**: \n   - Identify the owners of ABC Corp and XYZ Inc. through public information.\n\n8. **Consult HR**: \n   - Discuss the information collected to date and decide whether HR should participate in interviews.\n\n9. **Conduct Interviews**: \n   - Interview relevant individuals, including those from Finance, department heads, and the alleged employees (Joe Smith and Karen Jones).\n\n10. **Draft Investigation Report**: \n   - Prepare the investigation report with findings and recommendations, and discuss it with HR and the remediation committee.\n\nThis workplan is designed to ensure a thorough and confidential investigation in line with the company's policies and procedures." },
                    { "IV. Observations\n1. **Documentary Evidence:**  \n   - Summary of relevant documents reviewed and key information extracted.\n   \n2. **Interview Summaries:**  \n   - Synopses of interviews conducted, highlighting significant statements and insights.\n   \n3. **Compliance with Policies and Procedures:**  \n   - Assessment of adherence to company policies and procedures by the involved parties.\n\n", "1. **Documentary Evidence:**  \\n   - A detailed examination of Invoice 1592548 from ABC Corp revealed a $50,000 charge for \\\"consulting advice,\\\" a service that is not typically provided by the vendor. The invoice was approved by Karen Jones, which bypassed the standard approval process requiring Department Head authorization for amounts exceeding $40,000.  \\n   - Analysis of selection documents for XYZ Inc. indicated insufficient business justification for selecting the vendor over lower-priced bids. The justification labeled as \\\"best choice\\\" lacked a detailed explanation, which raised concerns about the objectivity of the selection process.\\n\\n2. **Interview Summaries:**  \\n   - Sam Walsh expressed unawareness regarding the specifics of the \\\"consulting advice\\\" service and acknowledged that the invoice approval did not adhere to the company’s established procedures.  \\n   - Joe Smith admitted to having a personal relationship with XYZ Inc., owned by his wife, and failed to disclose this conflict of interest. He could not provide a valid reason for instructing XYZ Inc. to increase an invoice by $5,000.  \\n   - Becky Hall recognized that the business justification for selecting XYZ Inc. was insufficient and committed to more diligently reviewing selection documents in the future.  \\n   - Mark Linder was surprised by the investigation’s findings, stating he had no prior knowledge of personal ties between company employees and the vendors or any issues with the vendors' performance.\\n\\n3. **Compliance with Policies and Procedures:**  \\n   - The investigation identified significant deviations from the Corporate Direct Procurement Policy and the Corporate Code of Conduct Policy. Notably, Joe Smith's failure to disclose his conflict of interest with XYZ Inc. and the improper approval of invoices highlighted gaps in adherence to and enforcement of these policies.\",\r\n    \" V. Conclusions\\n- Statement on whether the allegations could be or could not be substantiated.\\n\\n\": \"The investigation substantiated parts of the whistleblower's allegations, particularly concerning Joe Smith's undisclosed conflict of interest with XYZ Inc. and the irregular approval of invoices. Notably, Joe Smith could not provide a valid reason for instructing XYZ Inc. to increase an invoice by $5,000, and Becky Hall recognized the business justification for selecting XYZ Inc. as insufficient. There were significant deviations from the Corporate Direct Procurement Policy and the Corporate Code of Conduct Policy identified during the investigation. These findings suggest a breach of company policies and indicate potential ethical violations in the procurement process. Additionally, some financial records and communications may not have been fully accessible.\",\r\n    \" VI. Recommendations\\n- Specific actions proposed to address the findings, prevent future occurrences, and remediate any harm done.\\n\\n\": \"- Initiate a comprehensive financial audit focusing on all transactions and contracts involving ABC Corp and XYZ Inc. to uncover any further irregularities.\\n- Revise and reinforce the procurement and invoice approval processes, ensuring that all employees are aware of and adhere to the proper procedures.\\n- Implement targeted training sessions on ethics, conflict of interest, and procurement policies for all employees, with a focus on those involved in procurement and financial roles.\\n- Consider appropriate disciplinary actions against Joe Smith and Karen Jones, up to and including termination, for their violations of company policies." },
                    { "VI. Recommendations\n- Specific actions proposed to address the findings, prevent future occurrences, and remediate any harm done.\n\n", "- Initiate a comprehensive financial audit focusing on all transactions and contracts involving ABC Corp and XYZ Inc. to uncover any further irregularities.\n- Revise and reinforce the procurement and invoice approval processes, ensuring that all employees are aware of and adhere to the proper procedures.\n- Implement targeted training sessions on ethics, conflict of interest, and procurement policies for all employees, with a focus on those involved in procurement and financial roles.\n- Consider appropriate disciplinary actions against Joe Smith and Karen Jones, up to and including termination, for their violations of company policies." },
                    { "VII. Limitations\n- Any limitations encountered during the investigation that may affect the findings or conclusions.\n\n", "VII. Limitations  \nThe investigation faced limitations due to the reliance on self-reported information during interviews and the absence of direct evidence from the whistleblower. Additionally, some financial records and communications may not have been fully accessible." }
            };

            var dictionarytem = new Dictionary<string, string>
            {
                    { "IV. Observations\n1. **Documentary Evidence:**  \n   - Summary of relevant documents reviewed and key information extracted.\n   \n2. **Interview Summaries:**  \n   - Synopses of interviews conducted, highlighting significant statements and insights.\n   \n3. **Compliance with Policies and Procedures:**  \n   - Assessment of adherence to company policies and procedures by the involved parties.\n\n", "1. **Documentary Evidence:**  \\n   - A detailed examination of Invoice 1592548 from ABC Corp revealed a $50,000 charge for \\\"consulting advice,\\\" a service that is not typically provided by the vendor. The invoice was approved by Karen Jones, which bypassed the standard approval process requiring Department Head authorization for amounts exceeding $40,000.  \\n   - Analysis of selection documents for XYZ Inc. indicated insufficient business justification for selecting the vendor over lower-priced bids. The justification labeled as \\\"best choice\\\" lacked a detailed explanation, which raised concerns about the objectivity of the selection process.\\n\\n2. **Interview Summaries:**  \\n   - Sam Walsh expressed unawareness regarding the specifics of the \\\"consulting advice\\\" service and acknowledged that the invoice approval did not adhere to the company’s established procedures.  \\n   - Joe Smith admitted to having a personal relationship with XYZ Inc., owned by his wife, and failed to disclose this conflict of interest. He could not provide a valid reason for instructing XYZ Inc. to increase an invoice by $5,000.  \\n   - Becky Hall recognized that the business justification for selecting XYZ Inc. was insufficient and committed to more diligently reviewing selection documents in the future.  \\n   - Mark Linder was surprised by the investigation’s findings, stating he had no prior knowledge of personal ties between company employees and the vendors or any issues with the vendors' performance.\\n\\n3. **Compliance with Policies and Procedures:**  \\n   - The investigation identified significant deviations from the Corporate Direct Procurement Policy and the Corporate Code of Conduct Policy. Notably, Joe Smith's failure to disclose his conflict of interest with XYZ Inc. and the improper approval of invoices highlighted gaps in adherence to and enforcement of these policies.\",\r\n    \" V. Conclusions\\n- Statement on whether the allegations could be or could not be substantiated.\\n\\n\": \"The investigation substantiated parts of the whistleblower's allegations, particularly concerning Joe Smith's undisclosed conflict of interest with XYZ Inc. and the irregular approval of invoices. Notably, Joe Smith could not provide a valid reason for instructing XYZ Inc. to increase an invoice by $5,000, and Becky Hall recognized the business justification for selecting XYZ Inc. as insufficient. There were significant deviations from the Corporate Direct Procurement Policy and the Corporate Code of Conduct Policy identified during the investigation. These findings suggest a breach of company policies and indicate potential ethical violations in the procurement process. Additionally, some financial records and communications may not have been fully accessible.\",\r\n    \" VI. Recommendations\\n- Specific actions proposed to address the findings, prevent future occurrences, and remediate any harm done.\\n\\n\": \"- Initiate a comprehensive financial audit focusing on all transactions and contracts involving ABC Corp and XYZ Inc. to uncover any further irregularities.\\n- Revise and reinforce the procurement and invoice approval processes, ensuring that all employees are aware of and adhere to the proper procedures.\\n- Implement targeted training sessions on ethics, conflict of interest, and procurement policies for all employees, with a focus on those involved in procurement and financial roles.\\n- Consider appropriate disciplinary actions against Joe Smith and Karen Jones, up to and including termination, for their violations of company policies." },
            };

            ModifyWordDocument(filePath, dictionary);
            // ModifyWordDocument(filePath, dictionarytem);
            // ModifyWordDocument_(filePath, dictionarytem);
        }

        private static void ModifyWordDocument(string filePath, Dictionary<string, string> dictionary)
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
                {
                    Body body = wordDoc.MainDocumentPart.Document.Body;

                    foreach (var entry in dictionary)
                    {
                        var parts = entry.Key.Split(new[] { '-' }, 2, StringSplitOptions.None);

                        string searchText = CleanText(parts[1]); // Texto a buscar

                        Console.WriteLine($"searchText: {searchText}");
                        string insertText = entry.Value; // Texto a insertar

                        foreach (var paragraph in body.Elements<Paragraph>())
                        {
                            foreach (var text in paragraph.Descendants<Text>())
                            {
                                if (CleanText(text.Text.Trim()).Contains(searchText))
                                {
                                    // Console.WriteLine($"Found match for: {searchText}");

                                    AddBreakToParagraph(paragraph, 2); // Insertar dos saltos de línea
                                    AddTextToParagraph(paragraph, insertText); // Agregar texto
                                    continue;
                                } else {
                                    string paragraphText = CleanText(GetParagraphText(paragraph));
                                    if (CleanText(paragraphText.Trim()).Contains(searchText))
                                    {
                                        Console.WriteLine($"Found match for paragraphText: {searchText}");

                                        AddBreakToParagraph(paragraph, 2); // Insertar dos saltos de línea
                                        AddTextToParagraph(paragraph, insertText); // Agregar texto
                                        continue;
                                    } else {
                                        Console.WriteLine($"No match for: {searchText}");
                                        Console.WriteLine($"paragraphText: {paragraphText}");
                                        Console.WriteLine("\n\n");
                                    }
                                }
                            }
                        }
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

        private static void AddBreakToParagraph(Paragraph paragraph, int count)
        {
            for (int i = 0; i < count; i++)
            {
                paragraph.Append(new Run(new Break()));
            }
        }

        private static string CleanText(string text)
        {
            return new string(text
                .Where(c => !char.IsPunctuation(c)) // Eliminar puntuación
                .Select(c => char.ToLower(c)) // Convertir a minúsculas
                .ToArray())
                .Trim();
        }

        private static string GetParagraphText(Paragraph paragraph)
        {
            return string.Join("", paragraph.Descendants<Text>().Select(t => t.Text));
        }

    }
}
