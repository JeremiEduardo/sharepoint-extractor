import express from "express";
import axios from "axios";
import mammoth from "mammoth";

const app = express();
app.use(express.json());

// Función para listar todos los sites a los que el usuario tiene acceso
async function listAllSites(token) {
  const sites = [];
  let url = "https://graph.microsoft.com/v1.0/sites?search=*";

  while (url) {
    const res = await axios.get(url, { headers: { Authorization: token } });
    sites.push(...res.data.value);
    url = res.data["@odata.nextLink"] || null;
  }

  return sites;
}

// Función para listar todos los drives de un site
async function listDrives(siteId, token) {
  const res = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    { headers: { Authorization: token } }
  );
  return res.data.value;
}

// Función para listar todos los items de un drive (raíz y carpetas)
async function listAllDriveItems(driveId, token, parentId = "root") {
  const items = [];
  const queue = [parentId];

  while (queue.length > 0) {
    const current = queue.shift();
    const url =
      current === "root"
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`
        : `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${current}/children`;

    const res = await axios.get(url, { headers: { Authorization: token } });
    for (const item of res.data.value) {
      items.push(item);
      if (item.folder) queue.push(item.id);
    }
  }

  return items;
}

// Endpoint principal: buscar archivo por nombre y extraer texto
app.post("/search-and-extract", async (req, res) => {
  try {
    const { filename } = req.body;
    const token = req.headers.authorization;

    if (!token) return res.status(401).json({ error: "Falta token" });

    const sites = await listAllSites(token);

    for (const site of sites) {
      const drives = await listDrives(site.id, token);

      for (const drive of drives) {
        const items = await listAllDriveItems(drive.id, token);

        const file = items.find((i) => i.name === filename);
        if (file) {
          // Descargar archivo y extraer texto
          const fileResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${file.id}/content`,
            { headers: { Authorization: token }, responseType: "arraybuffer" }
          );

          const result = await mammoth.extractRawText({
            buffer: fileResponse.data
          });

          return res.json({
            site: site.webUrl,
            driveName: drive.name,
            text: result.value
          });
        }
      }
    }

    res.json({ error: "Archivo no encontrado en ningún site" });
  } catch (error) {
    console.error(error.response?.data || error.message);
    res.status(500).json({ error: "Error buscando o extrayendo archivo" });
  }
});

app.listen(process.env.PORT || 3000, () =>
  console.log("Servidor listo en puerto 3000")
);
