import express from 'express';
import axios from 'axios';
import dotenv from 'dotenv';
import open from 'open';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import exceljs from 'exceljs';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const port = 3000;
const excelFolder = path.join(__dirname, 'excel');

if (!fs.existsSync(excelFolder)) {
  fs.mkdirSync(excelFolder, { recursive: true });
}

let accessToken = null;

app.get('/', (req, res) => {
  res.send('<a href="/login">Login with Spotify</a>');
});

app.get('/login', (req, res) => {
  const authUrl = `https://accounts.spotify.com/authorize?response_type=code&client_id=${process.env.SPOTIFY_CLIENT_ID}&scope=playlist-read-private%20user-library-read&redirect_uri=${process.env.REDIRECT_URI}`;
  res.redirect(authUrl);
});

app.get('/callback', async (req, res) => {
  const code = req.query.code;

  try {
    const credentials = Buffer.from(`${process.env.SPOTIFY_CLIENT_ID}:${process.env.SPOTIFY_CLIENT_SECRET}`).toString('base64');

    const response = await axios.post(
      'https://accounts.spotify.com/api/token',
      new URLSearchParams({
        grant_type: 'authorization_code',
        code: code,
        redirect_uri: process.env.REDIRECT_URI
      }),
      {
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'Authorization': `Basic ${credentials}`
        }
      }
    );

    accessToken = response.data.access_token;
    res.send('Authentication successful! Now go to <a href="/playlists">/playlists</a>');
  } catch (error) {
    console.error('Error logging in:', error.response?.data || error.message);
    res.status(500).send('Error logging in');
  }
});

app.get('/playlists', async (req, res) => {
  if (!accessToken) {
    return res.redirect('/login');
  }

  try {
    const response = await axios.get('https://api.spotify.com/v1/me/playlists', {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    let html = '<h1>Your Playlists</h1><ul>';
    response.data.items.forEach(playlist => {
      html += `<li><a href="/download/${playlist.id}">${playlist.name}</a></li>`;
    });
    html += '</ul>';

    res.send(html);
  } catch (error) {
    console.error('Error fetching playlists:', error.response?.data || error.message);
    res.status(500).send('Error retrieving playlists');
  }
});

app.get('/download/:playlistId', async (req, res) => {
  if (!accessToken) {
    return res.redirect('/login');
  }

  try {
    const playlistId = req.params.playlistId;
    let allTracks = [];
    let nextUrl = `https://api.spotify.com/v1/playlists/${playlistId}/tracks?limit=100`;

    while (nextUrl) {
      const response = await axios.get(nextUrl, {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      allTracks = allTracks.concat(response.data.items);
      nextUrl = response.data.next;
    }

    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('Tracks');

    worksheet.columns = [
      { header: 'Track Name', key: 'track' },
      { header: 'Artist', key: 'artist' },
      { header: 'Album', key: 'album' }
    ];

    allTracks.forEach(item => {
      if (item.track) {
        worksheet.addRow({
          track: item.track.name,
          artist: item.track.artists.map(a => a.name).join(', '),
          album: item.track.album?.name || `${item.track.name} Song`
        });
      }
    });

    const filePath = path.join(excelFolder, `${playlistId}.xlsx`);
    await workbook.xlsx.writeFile(filePath);

    res.download(filePath, `${playlistId}.xlsx`);
  } catch (error) {
    console.error('Error downloading playlist:', error.response?.data || error.message);
    res.status(500).send('Error downloading playlist');
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
  // open(`http://localhost:${port}`);
});