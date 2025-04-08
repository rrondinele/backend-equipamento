const express = require('express');
const sql = require('mssql');
const cors = require('cors');
const ExcelJS = require('exceljs');
require('dotenv').config();

const app = express();
const port = process.env.PORT || 3001;

const corsOptions = {
  origin: [
    'http://localhost:3000',
    'https://query-equipamento.vercel.app'
  ],
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true
};
app.use(cors(corsOptions));
app.use(express.json());

const config = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER,
  database: process.env.DB_NAME,
  port: parseInt(process.env.DB_PORT),
  options: {
    encrypt: true,
    trustServerCertificate: false
  }
};

app.get('/', (req, res) => {
  res.send('Backend está funcionando!');
});

app.get('/api/data/:tableName', async (req, res) => {
  try {
    const pool = await sql.connect(config);
    const result = await pool.request()
      .query(`SELECT TOP 100 * FROM ${req.params.tableName}`);
    res.json(result.recordset);
  } catch (err) {
    console.error('Erro ao consultar o banco de dados:', err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/tables', async (req, res) => {
  try {
    const pool = await sql.connect(config);
    const result = await pool.request()
      .query("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'");
    res.json(result.recordset.map(item => item.TABLE_NAME));
  } catch (err) {
    console.error('Erro ao listar tabelas:', err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/equipamentos', async (req, res) => {
  try {
    const { dataInicial, dataFinal, equipamento, nota } = req.query;
    const pool = await sql.connect(config);
    let query = `
      SELECT TOP 20 
        [Instalação], [Nota], [Cliente], [Texto breve para o code],
        [Alavanca], CONVERT(VARCHAR, [Data Conclusão], 120) AS [Data Conclusão],
        [Equipamento Removido], [Material Removido], [Descrição Mat. Removido],
        [Status Equip. Removido], [Equipamento Instalado], [Material Instalado],
        [Descrição Mat. Instalado], [Status Equip. Instalado]
      FROM dbo.vw_equipe_removido
      WHERE 1=1
    `;

    if (dataInicial && dataFinal) {
      query += ` AND [Data Conclusão] BETWEEN @dataInicial AND @dataFinal`;
    }

    if (equipamento) {
      const equipamentos = equipamento.split(',').map(e => e.trim());
      if (equipamentos.length === 1) {
        query += ` AND [Equipamento Removido] = @equipamento`;
      } else {
        const paramsList = equipamentos.map((_, i) => `@equip${i}`).join(',');
        query += ` AND [Equipamento Removido] IN (${paramsList})`;
      }
    }

    if (nota) {
      const notas = nota.split(',').map(n => n.trim());
      if (notas.length === 1) {
        query += ` AND [Nota] = @nota`;
      } else {
        const notaParams = notas.map((_, i) => `@nota${i}`).join(',');
        query += ` AND [Nota] IN (${notaParams})`;
      }
    }
   

    query += ` ORDER BY [Data Conclusão] DESC`;
    const request = pool.request();

    if (dataInicial && dataFinal) {
      request.input('dataInicial', sql.Date, dataInicial);
      request.input('dataFinal', sql.Date, dataFinal);
    }

    if (equipamento) {
      const equipamentos = equipamento.split(',').map(e => e.trim());
      if (equipamentos.length === 1) {
        request.input('equipamento', sql.NVarChar, equipamentos[0]);
      } else {
        equipamentos.forEach((e, i) => {
          request.input(`equip${i}`, sql.NVarChar, e);
        });
      }
    }

    if (nota) {
      const notas = nota.split(',').map(n => n.trim());
      if (notas.length === 1) {
        request.input('nota', sql.NVarChar, notas[0]);
      } else {
        notas.forEach((n, i) => {
          request.input(`nota${i}`, sql.NVarChar, n);
        });
      }
    }    

    const result = await request.query(query);
    res.json(result.recordset);
  } catch (err) {
    console.error('Erro completo na consulta:', err);
    res.status(500).json({ error: 'Erro ao consultar equipamentos' });
  }
});

app.get('/api/equipamentos/count', async (req, res) => {
  try {
    const { dataInicial, dataFinal, equipamento } = req.query;
    const pool = await sql.connect(config);
    let query = `SELECT COUNT(*) AS count FROM dbo.vw_equipe_removido WHERE 1=1`;

    if (dataInicial && dataFinal) {
      query += ` AND [Data Conclusão] BETWEEN @dataInicial AND @dataFinal`;
    }

    if (equipamento) {
      const equipamentos = equipamento.split(',').map(e => e.trim());
      if (equipamentos.length === 1) {
        query += ` AND [Equipamento Removido] = @equipamento`;
      } else {
        const paramsList = equipamentos.map((_, i) => `@equip${i}`).join(',');
        query += ` AND [Equipamento Removido] IN (${paramsList})`;
      }
    }

    const request = pool.request();

    if (dataInicial && dataFinal) {
      request.input('dataInicial', sql.Date, dataInicial);
      request.input('dataFinal', sql.Date, dataFinal);
    }

    if (equipamento) {
      const equipamentos = equipamento.split(',').map(e => e.trim());
      if (equipamentos.length === 1) {
        request.input('equipamento', sql.NVarChar, equipamentos[0]);
      } else {
        equipamentos.forEach((e, i) => {
          request.input(`equip${i}`, sql.NVarChar, e);
        });
      }
    }

    const result = await request.query(query);
    res.json({ count: result.recordset[0].count });
  } catch (err) {
    console.error('Erro ao buscar contagem:', err);
    res.status(500).json({ error: 'Erro ao buscar contagem' });
  }
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});