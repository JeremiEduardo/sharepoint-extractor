import express from "express";
import axios from "axios";
import mammoth from "mammoth";

const app = express();
app.use(express.json());

app.post("/search-and-extract", async (req, res) => {
  try {
    const { query } = req.body;
    const token = req.headers.authorization;

    if (!token) {
      return res.status(401).json({ error: "Falta token" });
    }

    // 1. Buscar archivo en todos los sites
    const search = await axios.post(
      "https://graph.microsoft.com/v1.0/search/query",
      {
        requests: [
          {
            entityTypes: ["driveItem"],
            query: { queryString: query }
          }
        ]
      },
      {
        headers: { Authorization: token }
      }
    );

    const hits = search.data.value[0].hitsContainers[0].hits;

    if (!hits || hits.length === 0) {
      return res.json({ text: "No se encontró el archivo" });
    }

    const item = hits[0].resource;

    const driveId = item.parentReference.driveId;
    const itemId = item.id;

    // 2. Descargar archivo
    const file = await axios.get(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
      {
        headers: { Authorization: token },
        responseType: "arraybuffer"
      }
    );

    // 3. Convertir a texto
    const result = await mammoth.extractRawText({
      buffer: file.data
    });

    res.json({
      text: result.value
    });

  } catch (error) {
    console.error(error.message);
    res.status(500).json({ error: "Error completo" });
  }
});

app.listen(process.env.PORT || 3000);
