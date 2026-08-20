import express from 'express';
import path from 'path';
import fs from 'fs';
import axios from 'axios';
import * as XLSX from 'xlsx';
import { createServer as createViteServer } from 'vite';

const app = express();
// Azure App Service (and most cloud hosts) assign the port via process.env.PORT.
// Fall back to 3000 for local development.
const PORT = Number(process.env.PORT) || 3000;

app.use(express.json());

// Paths for persistent data files on the server (acting as simple flat-file databases)
const RECORDS_FILE = path.join(process.cwd(), 'mileage_records.json');
const SESSION_FILE = path.join(process.cwd(), 'onedrive_session.json');

// --- Helper Functions for Data Persistence ---

interface MileageRecord {
  id: string;
  consultantName: string;
  origin: string;
  destination: string;
  days: number;
  distanceKm: number;
  totalDistanceKm: number;
  date: string;
  isRoundTrip: boolean;
  syncedToOneDrive?: boolean;
}

interface OneDriveSession {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
  userName: string;
  userEmail: string;
}

function loadRecords(): MileageRecord[] {
  try {
    if (fs.existsSync(RECORDS_FILE)) {
      const data = fs.readFileSync(RECORDS_FILE, 'utf-8');
      return JSON.parse(data);
    }
  } catch (err) {
    console.error('Error loading mileage records from file:', err);
  }
  return [];
}

function saveRecords(records: MileageRecord[]) {
  try {
    fs.writeFileSync(RECORDS_FILE, JSON.stringify(records, null, 2), 'utf-8');
  } catch (err) {
    console.error('Error saving mileage records to file:', err);
  }
}

function loadSession(): OneDriveSession | null {
  try {
    if (fs.existsSync(SESSION_FILE)) {
      const data = fs.readFileSync(SESSION_FILE, 'utf-8');
      return JSON.parse(data);
    }
  } catch (err) {
    console.error('Error loading OneDrive session:', err);
  }
  return null;
}

function saveSession(session: OneDriveSession | null) {
  try {
    if (session) {
      fs.writeFileSync(SESSION_FILE, JSON.stringify(session, null, 2), 'utf-8');
    } else if (fs.existsSync(SESSION_FILE)) {
      fs.unlinkSync(SESSION_FILE);
    }
  } catch (err) {
    console.error('Error saving OneDrive session:', err);
  }
}

// --- Validation Helpers ---

function isValidDate(dateStr: string): boolean {
  if (!dateStr) return false;
  const parts = dateStr.split('/');
  if (parts.length !== 3) return false;
  const [day, month, year] = parts.map(Number);
  const d = new Date(year, month - 1, day);
  return d.getFullYear() === year && d.getMonth() === month - 1 && d.getDate() === day;
}

function validateMileageRecord(body: any): { valid: true; data: any } | { valid: false; error: string } {
  const {
    consultantName,
    origin,
    destination,
    days,
    distanceKm,
    totalDistanceKm,
    date,
    isRoundTrip,
  } = body;

  if (!consultantName || typeof consultantName !== 'string' || !consultantName.trim()) {
    return { valid: false, error: 'El nombre del consultor es obligatorio.' };
  }
  if (!origin || typeof origin !== 'string' || !origin.trim()) {
    return { valid: false, error: 'El origen es obligatorio.' };
  }
  if (!destination || typeof destination !== 'string' || !destination.trim()) {
    return { valid: false, error: 'El destino es obligatorio.' };
  }
  if (origin.trim().toLowerCase() === destination.trim().toLowerCase()) {
    return { valid: false, error: 'El origen y el destino deben ser diferentes.' };
  }
  if (!date || !isValidDate(date)) {
    return { valid: false, error: 'La fecha es inválida. Usa el formato DD/MM/YYYY.' };
  }
  const daysNum = Number(days);
  if (!Number.isFinite(daysNum) || daysNum < 1 || !Number.isInteger(daysNum)) {
    return { valid: false, error: 'El número de días debe ser un entero mayor o igual a 1.' };
  }
  const distanceNum = Number(distanceKm);
  if (!Number.isFinite(distanceNum) || distanceNum <= 0) {
    return { valid: false, error: 'La distancia debe ser un número positivo.' };
  }
  const totalNum = Number(totalDistanceKm);
  if (!Number.isFinite(totalNum) || totalNum < 0) {
    return { valid: false, error: 'El total de kilómetros es inválido.' };
  }

  return {
    valid: true,
    data: {
      consultantName: consultantName.trim(),
      origin: origin.trim(),
      destination: destination.trim(),
      days: daysNum,
      distanceKm: Math.round(distanceNum * 100) / 100,
      totalDistanceKm: Math.round(totalNum * 100) / 100,
      date,
      isRoundTrip: typeof isRoundTrip === 'boolean' ? isRoundTrip : true,
    },
  };
}

// --- Microsoft OAuth Variables ---
const getOauthCredentials = () => {
  return {
    clientId: process.env.MICROSOFT_CLIENT_ID || '',
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET || '',
  };
};

// Refresh Microsoft Graph Token if expired
async function refreshAccessTokenIfNeeded(): Promise<string | null> {
  const session = loadSession();
  if (!session) return null;

  const bufferTime = 5 * 60 * 1000; // 5 minutes before expiration
  if (Date.now() < session.expiresAt - bufferTime) {
    return session.accessToken;
  }

  const { clientId, clientSecret } = getOauthCredentials();
  if (!clientId || !clientSecret) {
    console.error('Missing Microsoft Client ID or Secret in environment.');
    return null;
  }

  try {
    console.log('Refreshing Microsoft Graph access token...');
    const params = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      refresh_token: session.refreshToken,
      grant_type: 'refresh_token',
      scope: 'offline_access Files.ReadWrite',
    });

    const response = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      params.toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { access_token, refresh_token, expires_in } = response.data;

    // Fetch user profile again to ensure we keep session fresh
    const meResponse = await axios.get('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${access_token}` },
    });

    const updatedSession: OneDriveSession = {
      accessToken: access_token,
      refreshToken: refresh_token || session.refreshToken,
      expiresAt: Date.now() + expires_in * 1000,
      userName: meResponse.data.displayName || 'Usuario Microsoft',
      userEmail: meResponse.data.mail || meResponse.data.userPrincipalName || '',
    };

    saveSession(updatedSession);
    console.log('Microsoft Graph token refreshed successfully.');
    return access_token;
  } catch (err: any) {
    console.error('Failed to refresh Microsoft Graph token:', err.response?.data || err.message);
    // If refresh token is invalid/revoked, we clear the session
    if (err.response?.status === 400) {
      saveSession(null);
    }
    return null;
  }
}

// Ensure the destination folder exists on OneDrive
async function ensureOneDriveFolder(token: string, folderName: string): Promise<void> {
  try {
    await axios.get(`https://graph.microsoft.com/v1.0/me/drive/root:/${folderName}`, {
      headers: { Authorization: `Bearer ${token}` },
    });
  } catch (err: any) {
    if (err.response?.status === 404) {
      console.log(`Creating OneDrive folder: ${folderName}`);
      await axios.post(
        'https://graph.microsoft.com/v1.0/me/drive/root/children',
        {
          name: folderName,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'rename',
        },
        { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' } }
      );
    } else {
      throw err;
    }
  }
}

// Helper to push a list of records to OneDrive Excel
async function syncRecordToOneDrive(record: MileageRecord): Promise<boolean> {
  const token = await refreshAccessTokenIfNeeded();
  if (!token) return false;

  let year = new Date().getFullYear();
  let monthStr = String(new Date().getMonth() + 1).padStart(2, '0');

  if (record.date) {
    const parts = record.date.split('/');
    if (parts.length === 3) {
      year = parseInt(parts[2], 10) || year;
      monthStr = parts[1].padStart(2, '0');
    }
  }

  const monthNames = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];
  const monthName = monthNames[parseInt(monthStr, 10) - 1] || 'Mes';
  const fileName = `Kilometraje_${year}_${monthStr}_${monthName}.xlsx`;
  const folderPath = 'GeoConsult_Kilometraje';
  const onedriveFilePath = `/${folderPath}/${fileName}`;

  try {
    await ensureOneDriveFolder(token, folderPath);
    let existingRecords: any[] = [];

    // 1. Check if the monthly Excel file already exists on OneDrive
    console.log(`Checking if file exists in OneDrive: ${onedriveFilePath}`);
    const checkUrl = `https://graph.microsoft.com/v1.0/me/drive/root:${onedriveFilePath}`;
    let fileExists = false;
    let downloadUrl = '';

    try {
      const metadataResponse = await axios.get(checkUrl, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (metadataResponse.data && metadataResponse.data['@microsoft.graph.downloadUrl']) {
        fileExists = true;
        downloadUrl = metadataResponse.data['@microsoft.graph.downloadUrl'];
      }
    } catch (err: any) {
      if (err.response?.status !== 404) {
        throw err;
      }
    }

    // 2. If it exists, download and parse existing Excel content
    if (fileExists && downloadUrl) {
      console.log('Downloading existing OneDrive Excel file to append...');
      const fileResponse = await axios.get(downloadUrl, { responseType: 'arraybuffer' });
      const workbook = XLSX.read(fileResponse.data, { type: 'buffer' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];

      existingRecords = XLSX.utils.sheet_to_json(worksheet);
      console.log(`Loaded ${existingRecords.length} existing records from OneDrive.`);
    }

    // 3. Prepare the new row
    const newRow = {
      'Consultor': record.consultantName,
      'Origen (Domicilio)': record.origin,
      'Destino (Planta/Cliente)': record.destination,
      'Días': record.days,
      'Ida y Vuelta': (record.isRoundTrip ?? true) ? 'Sí' : 'No',
      'Distancia Trayecto (Km)': record.distanceKm,
      'Total Trayecto (Km)': Number((record.distanceKm * ((record.isRoundTrip ?? true) ? 2 : 1)).toFixed(2)),
      'Total Acumulado (Km)': record.totalDistanceKm,
      'Fecha Registro': record.date,
      'ID Registro': record.id,
    };

    // Append new row and prevent duplicates by checking ID
    const duplicateIndex = existingRecords.findIndex(r => r['ID Registro'] === record.id);
    if (duplicateIndex >= 0) {
      existingRecords[duplicateIndex] = newRow;
    } else {
      existingRecords.push(newRow);
    }

    // 4. Generate new Excel workbook
    const newWorkbook = XLSX.utils.book_new();

    // Detailed list sheet
    const detailSheet = XLSX.utils.json_to_sheet(existingRecords);
    XLSX.utils.book_append_sheet(newWorkbook, detailSheet, 'Detalle de Trayectos');

    // Create a dynamic summary sheet: grouped by Consultant
    const summaryMap = existingRecords.reduce((acc, r) => {
      const name = r['Consultor'];
      if (!acc[name]) {
        acc[name] = {
          'Consultor': name,
          'Total Viajes': 0,
          'Total Días': 0,
          'Km Totales Acumulados': 0
        };
      }
      acc[name]['Total Viajes'] += 1;
      acc[name]['Total Días'] += (Number(r['Días']) || 0);
      acc[name]['Km Totales Acumulados'] += (Number(r['Total Acumulado (Km)']) || 0);
      return acc;
    }, {} as Record<string, any>);

    const summarySheet = XLSX.utils.json_to_sheet(Object.values(summaryMap));
    XLSX.utils.book_append_sheet(newWorkbook, summarySheet, 'Resumen por Consultor');

    // Write workbook to buffer
    const excelBuffer = XLSX.write(newWorkbook, { type: 'buffer', bookType: 'xlsx' });

    // 5. Upload updated Excel workbook back to OneDrive
    console.log(`Uploading updated Excel to OneDrive path: ${onedriveFilePath}`);
    const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:${onedriveFilePath}:/content`;

    await axios.put(uploadUrl, excelBuffer, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      },
    });

    console.log('OneDrive synchronization completed successfully.');
    return true;
  } catch (err: any) {
    console.error('Error synchronizing Excel to OneDrive:', err.response?.data || err.message);
    return false;
  }
}

// --- API Router ---

// Get local server status & connected OneDrive account details
app.get('/api/auth/status', async (req, res) => {
  const session = loadSession();
  const credentials = getOauthCredentials();
  
  if (!session) {
    return res.json({
      connected: false,
      isConfigured: !!(credentials.clientId && credentials.clientSecret),
    });
  }

  const token = await refreshAccessTokenIfNeeded();
  if (!token) {
    return res.json({
      connected: false,
      isConfigured: !!(credentials.clientId && credentials.clientSecret),
    });
  }

  // Get fresh session to ensure latest data
  const currentSession = loadSession();
  res.json({
    connected: true,
    isConfigured: !!(credentials.clientId && credentials.clientSecret),
    userName: currentSession?.userName || 'Usuario Microsoft',
    userEmail: currentSession?.userEmail || '',
  });
});

// Construct & return OneDrive OAuth start URL
app.get('/api/auth/microsoft/url', (req, res) => {
  const { clientId } = getOauthCredentials();
  
  if (!clientId) {
    return res.status(500).json({ error: 'La clave de Cliente de Microsoft (MICROSOFT_CLIENT_ID) no está configurada.' });
  }

  // Self-referential absolute URL for callback
  const appUrl = process.env.APP_URL || `${req.protocol}://${req.get('host')}`;
  const redirectUri = `${appUrl}/api/auth/microsoft/callback`;

  const params = new URLSearchParams({
    client_id: clientId,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: 'query',
    scope: 'offline_access Files.ReadWrite',
    state: 'geoconsult_onedrive_auth',
  });

  const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?${params.toString()}`;
  res.json({ url: authUrl });
});

// Microsoft OAuth callback handler
app.get(['/api/auth/microsoft/callback', '/api/auth/microsoft/callback/'], async (req, res) => {
  const { code, error, error_description } = req.query;

  if (error) {
    console.error('Microsoft OAuth error returned:', error, error_description);
    return res.send(`
      <html>
        <body style="font-family: sans-serif; text-align: center; padding: 40px; color: #334155;">
          <h2 style="color: #ef4444;">Error de Autenticación</h2>
          <p>${error_description || error}</p>
          <button onclick="window.close()" style="background: #0f172a; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: 600;">Cerrar ventana</button>
        </body>
      </html>
    `);
  }

  const { clientId, clientSecret } = getOauthCredentials();
  const appUrl = process.env.APP_URL || `${req.protocol}://${req.get('host')}`;
  const redirectUri = `${appUrl}/api/auth/microsoft/callback`;

  try {
    const params = new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      code: String(code),
      redirect_uri: redirectUri,
      grant_type: 'authorization_code',
      scope: 'offline_access Files.ReadWrite',
    });

    const tokenResponse = await axios.post(
      'https://login.microsoftonline.com/common/oauth2/v2.0/token',
      params.toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { access_token, refresh_token, expires_in } = tokenResponse.data;

    // Fetch user details from Microsoft Graph
    const meResponse = await axios.get('https://graph.microsoft.com/v1.0/me', {
      headers: { Authorization: `Bearer ${access_token}` },
    });

    const sessionData: OneDriveSession = {
      accessToken: access_token,
      refreshToken: refresh_token,
      expiresAt: Date.now() + expires_in * 1000,
      userName: meResponse.data.displayName || 'Usuario Microsoft',
      userEmail: meResponse.data.mail || meResponse.data.userPrincipalName || '',
    };

    saveSession(sessionData);

    // Send postMessage and close popup
    res.send(`
      <html>
        <body style="font-family: sans-serif; text-align: center; padding: 40px; color: #334155;">
          <h2 style="color: #10b981;">¡Conectado con éxito!</h2>
          <p>Se ha vinculado tu cuenta de OneDrive.</p>
          <p style="font-size: 13px; color: #94a3b8;">Esta ventana se cerrará automáticamente en unos segundos.</p>
          <script>
            if (window.opener) {
              window.opener.postMessage({ type: 'MS_AUTH_SUCCESS' }, '*');
            }
            setTimeout(function() {
              window.close();
            }, 1500);
          </script>
        </body>
      </html>
    `);
  } catch (err: any) {
    console.error('Error exchanging OAuth code:', err.response?.data || err.message);
    res.status(500).send(`
      <html>
        <body style="font-family: sans-serif; text-align: center; padding: 40px; color: #334155;">
          <h2 style="color: #ef4444;">Error de Intercambio de Token</h2>
          <p>${err.response?.data?.error_description || err.message}</p>
          <button onclick="window.close()" style="background: #0f172a; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: 600;">Cerrar ventana</button>
        </body>
      </html>
    `);
  }
});

// Logout from OneDrive
app.post('/api/auth/logout', (req, res) => {
  saveSession(null);
  res.json({ success: true });
});

// Fetch all local cached records
app.get('/api/records', (req, res) => {
  const records = loadRecords();
  res.json(records);
});

// Create a new record and sync to OneDrive if connected
app.post('/api/records', async (req, res) => {
  const validation = validateMileageRecord(req.body);
  if (!validation.valid) {
    return res.status(400).json({ error: (validation as { error: string }).error });
  }

  const recordId = req.body.id || Math.random().toString(36).substring(2, 11);
  const newRecord: MileageRecord = {
    ...validation.data,
    id: recordId,
    syncedToOneDrive: false,
  };

  const records = loadRecords();
  records.unshift(newRecord);
  saveRecords(records);

  // Sync to OneDrive if possible
  const isConnected = !!(await refreshAccessTokenIfNeeded());
  if (isConnected) {
    const syncSuccess = await syncRecordToOneDrive(newRecord);
    if (syncSuccess) {
      newRecord.syncedToOneDrive = true;
      const updatedRecords = loadRecords();
      const updatedIdx = updatedRecords.findIndex(r => r.id === recordId);
      if (updatedIdx >= 0) {
        updatedRecords[updatedIdx].syncedToOneDrive = true;
        saveRecords(updatedRecords);
      }
    }
  }

  res.json(newRecord);
});

// Update an existing record and re-sync to OneDrive if connected
app.put('/api/records/:id', async (req, res) => {
  const { id } = req.params;
  const validation = validateMileageRecord(req.body);
  if (!validation.valid) {
    return res.status(400).json({ error: (validation as { error: string }).error });
  }

  const records = loadRecords();
  const index = records.findIndex(r => r.id === id);
  if (index === -1) {
    return res.status(404).json({ error: 'Registro no encontrado' });
  }

  const updatedRecord: MileageRecord = {
    ...validation.data,
    id,
    syncedToOneDrive: false,
  };

  records[index] = updatedRecord;
  saveRecords(records);

  const isConnected = !!(await refreshAccessTokenIfNeeded());
  if (isConnected) {
    const syncSuccess = await syncRecordToOneDrive(updatedRecord);
    if (syncSuccess) {
      updatedRecord.syncedToOneDrive = true;
      const reloaded = loadRecords();
      const reloadedIdx = reloaded.findIndex(r => r.id === id);
      if (reloadedIdx >= 0) {
        reloaded[reloadedIdx].syncedToOneDrive = true;
        saveRecords(reloaded);
      }
    }
  }

  res.json(updatedRecord);
});

// Bulk sync pending local records to OneDrive
app.post('/api/sync', async (req, res) => {
  const isConnected = !!(await refreshAccessTokenIfNeeded());
  if (!isConnected) {
    return res.status(401).json({ error: 'OneDrive no está conectado' });
  }

  const records = loadRecords();
  let syncCount = 0;

  for (const record of records) {
    const success = await syncRecordToOneDrive(record);
    if (success) {
      record.syncedToOneDrive = true;
      syncCount++;
    }
  }

  saveRecords(records);
  res.json({ success: true, syncedCount: syncCount, totalCount: records.length });
});

// Delete a record
app.delete('/api/records/:id', (req, res) => {
  const { id } = req.params;
  const records = loadRecords();
  const filtered = records.filter(r => r.id !== id);
  saveRecords(filtered);
  res.json({ success: true });
});

// --- Vite Middleware / Frontend Serving ---

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on port ${PORT}`);
  });
}

startServer();
