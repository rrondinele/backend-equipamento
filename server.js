// server.js otimizado
const express = require('express');
const sql = require('mssql');
const cors = require('cors');
const XLSX = require('xlsx');
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
    trustServerCertificate: true,
  },
  connectionTimeout: 30000,
  requestTimeout: 30000
};

const aplicarFiltros = (query, filtros, campos) => {
  campos.forEach(({ nome, coluna }) => {
    const valor = filtros[nome];
    if (valor && valor !== 'Todos') {
      const lista = valor.split(',').map(v => v.trim()).filter(v => v && v !== 'Todos');
      if (lista.length === 0) return; // ignora filtro vazio

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

const consultarEquipamentos = async (filtros, isCount = false, limite = null) => {
  const pool = await sql.connect(config);

  let topClause = '';
  if (!isCount && limite) {
    topClause = `TOP ${limite}`;
  }

  const query = {
    sql: isCount
      ? 'SELECT COUNT(*) AS count FROM dbo.vw_equipe_removido WHERE 1=1'
      : `
        SELECT ${topClause}
          [Instalação], [Nota], [Cliente], [Texto breve para o code], [Alavanca],
          CONVERT(VARCHAR, [Data Conclusão], 120) AS [Data Conclusão],
          [Equipamento Removido], [Material Removido], [Descrição Mat. Removido], [Status Equip. Removido],
          [Equipamento Instalado], [Material Instalado], [Descrição Mat. Instalado], [Status Equip. Instalado]
        FROM dbo.vw_equipe_removido WHERE 1=1`,
    inputs: []
  };

  if (filtros.dataInicial && filtros.dataFinal) {
    query.sql += ' AND [Data Conclusão] BETWEEN @dataInicial AND @dataFinal';
    query.inputs.push({ key: 'dataInicial', value: filtros.dataInicial, type: sql.Date });
    query.inputs.push({ key: 'dataFinal', value: filtros.dataFinal, type: sql.Date });
  }

  aplicarFiltros(query, filtros, [
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
    const data = await consultarEquipamentos(req.query, false, 20); // <- aqui limitando para 20
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

app.get('/api/equipamentos/export', async (req, res) => {
  try {
    const data = await consultarEquipamentos(req.query, false);
    res.json(data); // ← envia os dados em JSON puro
  } catch (error) {
    console.error("Erro ao exportar:", error);
    res.status(500).json({ error: "Erro ao exportar dados" });
  }
});

// OFS_Materiais
const consultarMateriais = async (filtros, isCount = false, limite = null) => {
  const pool = await sql.connect(config);

  let topClause = '';
  if (!isCount && limite) { // Só aplica TOP se limite for fornecido
    topClause = `TOP ${limite}`;
  }

  const query = {
    sql: isCount
      ? 'SELECT COUNT(*) AS count FROM dbo.vw_ofs_material WHERE 1=1'
      : `
        SELECT ${topClause}
          [Data], [Nota], [Texto Breve], [Acao],
          [Status do Usuário], [Tipo de nota], [Instalação],
          [Zona], [Lote], [Descricao], [Quantidade], [Serial], [Base Operacional]
        FROM dbo.vw_ofs_material WHERE 1=1`,
    inputs: []
  };

  if (filtros.dataInicial && filtros.dataFinal) {
    query.sql += ' AND [Data] BETWEEN @dataInicial AND @dataFinal';
    query.inputs.push({ key: 'dataInicial', value: filtros.dataInicial, type: sql.Date });
    query.inputs.push({ key: 'dataFinal', value: filtros.dataFinal, type: sql.Date });
  }

  aplicarFiltros(query, filtros, [
    { nome: 'nota', coluna: 'Nota' },
    { nome: 'equipamento', coluna: 'Serial' }, 
    { nome: 'status', coluna: 'Acao' }
  ]);

  const request = pool.request();
  query.inputs.forEach(({ key, value, type }) => {
    request.input(key, type || sql.NVarChar, value);
  });

  const result = await request.query(query.sql);
  return result.recordset;
};

app.get('/api/materiais', async (req, res) => {
  try {
    const data = await consultarMateriais(req.query, false, 20);
    res.json(data);
  } catch (err) {
    console.error('Erro ao consultar materiais:', err);
    res.status(500).json({ error: 'Erro ao consultar materiais' });
  }
});

app.get('/api/materiais/export', async (req, res) => {
  try {
    // Remove qualquer limite para a exportação
    const data = await consultarMateriais(req.query, false); // Sem o parâmetro limite
    res.json(data);
  } catch (err) {
    console.error('Erro ao exportar materiais:', err);
    res.status(500).json({ error: 'Erro ao exportar materiais' });
  }
});

app.get('/api/materiais/count', async (req, res) => {
  try {
    const data = await consultarMateriais(req.query, true);
    res.json({ count: data[0]?.count || 0 });
  } catch (err) {
    console.error('Erro ao contar materiais:', err);
    res.status(500).json({ error: 'Erro ao contar materiais' });
  }
});

app.listen(port, () => console.log(`Servidor rodando na porta ${port}`));