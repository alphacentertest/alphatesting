const crypto = require('crypto');
const blob = require('@vercel/blob');
const fetch = async (...args) => {
  const { default: fetch } = await import('node-fetch');
  return fetch(...args);
};
const Store = require('express-session').Store;

if (!process.env.SESSION_SECRET) {
  throw new Error('SESSION_SECRET environment variable is not set');
}

if (!process.env.BLOB_READ_WRITE_TOKEN) {
  throw new Error('BLOB_READ_WRITE_TOKEN environment variable is not set');
}

const BASE_URL = 'https://qqeygegbb01p35fz.public.blob.vercel-storage.com'; // Your Vercel Blob Store base URL

const algorithm = 'aes-256-cbc';
const key = crypto.scryptSync(process.env.SESSION_SECRET, 'salt', 32);

class VercelBlobStore extends Store {
  constructor(options = {}) {
    super();
    this.prefix = options.prefix || 'sess/'; // Use a path prefix
  }

  async get(sid, callback) {
    try {
      const blobName = `${this.prefix}${sid}`;
      // Use blob.head to check if the blob exists
      const headResponse = await blob.head(blobName, { token: process.env.BLOB_READ_WRITE_TOKEN });
      if (!headResponse) {
        return callback(null, null); // Blob not found
      }
      const url = blob.getDownloadUrl(blobName, process.env.BLOB_READ_WRITE_TOKEN);
      console.log('Generated download URL:', url);
      const response = await fetch(url);
      if (!response.ok) {
        if (response.status === 404) {
          return callback(null, null); // Blob not found
        }
        throw new Error(`Failed to fetch blob: ${response.statusText}`);
      }
      const encryptedData = await response.text();
      const [ivHex, encrypted] = encryptedData.split(':');
      const decipher = crypto.createDecipheriv(algorithm, key, Buffer.from(ivHex, 'hex'));
      let decrypted = decipher.update(encrypted, 'hex', 'utf8');
      decrypted += decipher.final('utf8');
      callback(null, JSON.parse(decrypted));
    } catch (error) {
      console.error('Error retrieving session from Blob Store:', error);
      callback(error);
    }
  }

  async set(sid, session, callback) {
    try {
      const blobName = `${this.prefix}${sid}`;
      const iv = crypto.randomBytes(16);
      const cipher = crypto.createCipheriv(algorithm, key, iv);
      let encrypted = cipher.update(JSON.stringify(session), 'utf8', 'hex');
      encrypted += cipher.final('hex');
      const encryptedData = `${iv.toString('hex')}:${encrypted}`;
      await blob.put(blobName, encryptedData, { 
        access: 'public', 
        token: process.env.BLOB_READ_WRITE_TOKEN,
        allowOverwrite: true // Allow overwriting existing blobs
      });
      callback(null);
    } catch (error) {
      console.error('Error storing session in Blob Store:', error);
      callback(error);
    }
  }

  async destroy(sid, callback) {
    try {
      const blobName = `${this.prefix}${sid}`;
      await blob.del(blobName, { token: process.env.BLOB_READ_WRITE_TOKEN });
      callback(null);
    } catch (error) {
      console.error('Error deleting session from Blob Store:', error);
      callback(error);
    }
  }

  async touch(sid, session, callback) {
    this.set(sid, session, callback);
  }
}

module.exports = VercelBlobStore;
