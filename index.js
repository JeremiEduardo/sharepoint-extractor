import express from "express";
import axios from "axios";
import mammoth from "mammoth";

const app = express();
app.use(express.json());

app.post("/extract-text", async (req, res) => {
  try {
    const { driveId, itemId } = req.body;

    // Aquí recibirás el token desde la acción
    const token = req.headers.authorization;

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
    res.status(500).json({
      error: "Error extrayendo texto"
    });
  }
});

app.listen(process.env.PORT || 3000);