import express from "express";
import axios from "axios";
import mammoth from "mammoth";

const app = express();
app.use(express.json());

app.get("/", (req, res) => {
  res.send("SharePoint Extractor activo 🚀");
});

app.post("/extract-text", async (req, res) => {
  try {
    const { driveId, itemId } = req.body;

    if (!driveId || !itemId) {
      return res.status(400).json({
        error: "driveId e itemId son requeridos"
      });
    }

    const token = req.headers.authorization;

    if (!token) {
      return res.status(401).json({
        error: "Falta Authorization header"
      });
    }

    // Descargar archivo desde Microsoft Graph
    const fileResponse = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
      {
        headers: {
          Authorization: token
        },
        responseType: "arraybuffer"
      }
    );

    // Convertir DOCX a texto
    const result = await mammoth.extractRawText({
      buffer: fileResponse.data
    });

    res.json({
      text: result.value
    });

  } catch (error) {
    console.error("Error:", error.response?.data || error.message);

    res.status(500).json({
      error: "Error extrayendo texto"
    });
  }
});

const PORT = process.env.PORT;

app.listen(PORT, () => {
  console.log(`Servidor corriendo en puerto ${PORT}`);
});
