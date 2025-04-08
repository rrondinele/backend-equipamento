// server.js otimizado
const express = require('express');
const sql = require('mssql');
const cors = require('cors');
const XLSX = require("xlsx");
require('dotenv').config();

const app = express();
const port = process.env.PORT || 3001;

const corsOptions = {
  origin: ['http://localhost:3000', 'https://query-equipamento.vercel.app'],
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

const aplicarFiltros = (query, params, filtros, campos) => {
  campos.forEach(({ nome, coluna }) => {
    const valor = filtros[nome];
    if (valor) {
      const lista = valor.split(',').map(v => v.trim());
      if (lista.length === 1) {
        query.sql += ` AND [${coluna}] = @${nome}`;
        query.inputs.push({ key: nome, value: lista[0] });
      } else {
        const keys = lista.map((_, i) => `@${nome}${i}`).join(',');
        query.sql += ` AND [${coluna}] IN (${keys})`;
        lista.forEach((v, i) => {
          query.inputs.push({ key: `${nome}${i}`, value: v });
        });
      }
    }
  });
};

app.get('/', (req, res) => res.send('Backend está funcionando!'));

app.get('/api/data/:tableName', async (req, res) => {
  try {
    const pool = await sql.connect(config);
    const result = await pool.request().query(`SELECT TOP 100 * FROM ${req.params.tableName}`);
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

const consultarEquipamentos = async (filtros, isCount = false) => {
  const pool = await sql.connect(config);
  const query = {
    sql: isCount ? 'SELECT COUNT(*) AS count FROM dbo.vw_equipe_removido WHERE 1=1' : `
      SELECT TOP 20 
        [Instalação], [Nota], [Cliente], [Texto breve para o code],
        [Alavanca], CONVERT(VARCHAR, [Data Conclusão], 120) AS [Data Conclusão],
        [Equipamento Removido], [Material Removido], [Descrição Mat. Removido],
        [Status Equip. Removido], [Equipamento Instalado], [Material Instalado],
        [Descrição Mat. Instalado], [Status Equip. Instalado]
      FROM dbo.vw_equipe_removido WHERE 1=1`,
    inputs: []
  };

  if (filtros.dataInicial && filtros.dataFinal) {
    query.sql += ' AND [Data Conclusão] BETWEEN @dataInicial AND @dataFinal';
    query.inputs.push({ key: 'dataInicial', value: filtros.dataInicial, type: sql.Date });
    query.inputs.push({ key: 'dataFinal', value: filtros.dataFinal, type: sql.Date });
  }

  aplicarFiltros(query, filtros, filtros, [
    { nome: 'equipamento', coluna: 'Equipamento Removido' },
    { nome: 'nota', coluna: 'Nota' }
  ]);

  const request = pool.request();
  query.inputs.forEach(({ key, value, type }) => {
    request.input(key, type || sql.NVarChar, value);
  });

  const result = await request.query(query.sql);
  return result.recordset;
};

app.get('/api/equipamentos', async (req, res) => {
  try {
    const data = await consultarEquipamentos(req.query, false);
    res.json(data);
  } catch (err) {
    console.error('Erro completo na consulta:', err);
    res.status(500).json({ error: 'Erro ao consultar equipamentos' });
  }
});

app.get('/api/equipamentos/count', async (req, res) => {
  try {
    const data = await consultarEquipamentos(req.query, true);
    res.json({ count: data[0]?.count || 0 });
  } catch (err) {
    console.error('Erro ao buscar contagem:', err);
    res.status(500).json({ error: 'Erro ao buscar contagem' });
  }
});

app.get('/api/equipamentos/ultima-data', async (req, res) => {
  try {
    const pool = await sql.connect(config);
    const result = await pool.request().query('SELECT MAX([Data Conclusão]) AS ultimaData FROM dbo.vw_equipe_removido');
    res.json({ ultimaData: result.recordset[0].ultimaData });
  } catch (err) {
    console.error('Erro ao obter última data:', err);
    res.status(500).json({ error: 'Erro ao obter última data' });
  }
});

app.get("/api/equipamentos/export", async (req, res) => {
  try {
    const { data_atividade, equipamento } = req.query;
    let filtered = [...dados];

    if (data_atividade) {
      filtered = filtered.filter((item) => item.data_atividade === data_atividade);
    }

    if (equipamento) {
      filtered = filtered.filter((item) => item.equipe === equipamento);
    }

    const worksheet = XLSX.utils.json_to_sheet(filtered);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dados");

    const buffer = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });

    res.setHeader("Content-Disposition", "attachment; filename=exportacao.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    return res.send(buffer);
  } catch (error) {
    console.error("Erro ao exportar:", error);
    res.status(500).json({ error: "Erro ao exportar dados" });
  }
});

app.listen(port, () => console.log(`Servidor rodando na porta ${port}`));
