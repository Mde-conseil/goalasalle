const { ClientSecretCredential } = require("@azure/identity");
const fetch = require("node-fetch");

const graphScope = "https://graph.microsoft.com/.default";

module.exports = async function (context, req) {
  try {
    const body = req.body || {};

    // Vérif minimum
    const required = ["Sexe", "Age", "Lieu", "SousLieu", "Motivation"];
    for (const field of required) {
      if (!body[field]) {
        context.res = {
          status: 400,
          body: `Champ manquant : ${field}`
        };
        return;
      }
    }

    // Authentification Azure AD (app en mode "client credentials")
    const credential = new ClientSecretCredential(
      process.env.TENANT_ID,
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET
    );

    const token = await credential.getToken(graphScope);
    if (!token || !token.token) {
      throw new Error("Impossible d’obtenir un token Graph");
    }

    const siteId = process.env.SITE_ID;   // ex : "contoso.sharepoint.com,...."
    const listId = process.env.LIST_ID;   // ID ou nom de ta liste

    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;

    // Mappage des champs vers la liste SharePoint
    const payload = {
      fields: {
        Sexe: body.Sexe,
        Age: Number(body.Age),
        Lieu: body.Lieu,
        SousLieu: body.SousLieu,
        Motivation: Number(body.Motivation)
      }
    };

    const graphRes = await fetch(graphUrl, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token.token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    if (!graphRes.ok) {
      const errorText = await graphRes.text();
      context.log("Graph error:", errorText);
      context.res = {
        status: graphRes.status,
        body: errorText
      };
      return;
    }

    context.res = {
      status: 200,
      body: "OK"
    };
  } catch (err) {
    context.log("Function error:", err);
    context.res = {
      status: 500,
      body: err.message || "Erreur interne"
    };
  }
};
