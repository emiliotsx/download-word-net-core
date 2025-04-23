# EJEMPLO DE POST

```
const generateWordDocument = async () => {
  const url = "http://localhost:5150/api/WordGenerator/generate"; // Cambia la URL según tu configuración
  const data = {
    "Case ID": "12345",
    "Report Date": "2025-04-21",
    "Report Issued by": "John Doe",
    "Summary": "This is a test report generated from a dictionary input.",
    "Observations": "No major issues detected.",
    "Conclusions": "Proceed with the current plan."
  }

  try {
    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(data)
    });

    if (!response.ok) {
      console.log('response', response)
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const blob = await response.blob();
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = "document.docx"; // Nombre del archivo descargado
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(downloadUrl);
  } catch (error) {
    console.log("Error:", error);
    console.error("Error generating Word document:", error);
  }
}

```