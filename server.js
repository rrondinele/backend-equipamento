const express = require('express');
const sql = require('mssql');
const cors = require('cors');
const mssql = require('mssql');
const ExcelJS = require('exceljs');
require('dotenv').config();

// Cria a aplicação Express
const app = express();
const port = process.env.PORT || 3001;

const corsOptions = {
    origin: [
      'http://localhost:3000',
      'https://query-equipamento.vercel.app' // Remova a barra no final
    ],
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: true
  };
  app.use(cors(corsOptions));

// Outros middlewares
app.use(express.json());

// Configuração do banco de dados
const config = {
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  server: process.env.DB_SERVER,
  database: process.env.DB_NAME,
  port: parseInt(process.env.DB_PORT),
  options: {
    encrypt: true, // Para Azure SQL
    trustServerCertificate: false
  }
};

// Rota de teste simples
app.get('/', (req, res) => {
  res.send('Backend está funcionando!');
});

// Rota para obter dados de uma tabela
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

// Rota para listar tabelas disponíveis
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


// Rota para consulta específica de equipamentos - Versão Corrigida
app.get('/api/equipamentos', async (req, res) => {
  try {
    const { dataInicial, dataFinal, equipamento } = req.query;
    
    const pool = await sql.connect(config);
    
    // Query corrigida com sintaxe válida
    let query = `
          SELECT TOP 20 
            [Instalação],
            [Nota],
            [Cliente],
            [Texto breve para o code],
            [Alavanca],
            CONVERT(VARCHAR, [Data Conclusão], 120) AS [Data Conclusão],
            [Equipamento Removido],
            [Material Removido],
            [Descrição Mat. Removido],
            [Status Equip. Removido],
            [Equipamento Instalado],
            [Material Instalado],
            [Descrição Mat. Instalado],
            [Status Equip. Instalado]
          FROM dbo.vw_equipe_removido  
          WHERE 1=1
    `;

    // Filtro por datas
    if (dataInicial && dataFinal) {
      query += ` AND [Data Conclusão] BETWEEN @dataInicial AND @dataFinal`;
    }

    // Filtro por equipamento
    if (equipamento) {
      const equipamentos = equipamento.split(',').map(e => e.trim());
      if (equipamentos.length === 1) {
        query += ` AND [Equipamento Removido] = @equipamento`;
      } else {
        // Versão mais segura usando parâmetros
        const paramsList = equipamentos.map((_, i) => `@equip${i}`).join(',');
        query += ` AND [Equipamento Removido] IN (${paramsList})`;
      }
    }

    query += ` ORDER BY [Data Conclusão] DESC`;

    const request = pool.request();
    
    // Adiciona parâmetros de data
    if (dataInicial && dataFinal) {
      request.input('dataInicial', sql.Date, dataInicial);
      request.input('dataFinal', sql.Date, dataFinal);
    }
    
    // Adiciona parâmetros de equipamento
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

    console.log('Executando query:', query); // Log para diagnóstico
    const result = await request.query(query);
    
    res.json(result.recordset);
    
  } catch (err) {
    console.error('Erro completo na consulta:', {
      message: err.message,
      stack: err.stack,
      originalError: err.originalError,
      query: req.query
    });
    
    res.status(500).json({ 
      error: 'Erro ao consultar equipamentos',
      details: process.env.NODE_ENV === 'development' ? err.message : undefined
    });
  }
});

// Rota para contar equipamentos com filtros
app.get('/api/equipamentos/count', async (req, res) => {
  try {
    await mssql.connect(dbConfig);

    const { equipamento, dataInicial, dataFinal } = req.query;
    const conditions = [];

    if (equipamento) {
      const equipamentos = equipamento.split(',').map(e => `'${e.trim()}'`).join(',');
      conditions.push(`[Equipamento Removido] IN (${equipamentos})`);
    }

    if (dataInicial && dataFinal) {
      conditions.push(`[Data Conclusão] BETWEEN '${dataInicial}' AND '${dataFinal}'`);
    }

    const whereClause = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';

    const query = `SELECT COUNT(*) AS count FROM [dbo].[equipamentos] ${whereClause}`;
    const result = await mssql.query(query);
    res.json({ count: result.recordset[0].count });
  } catch (err) {
    console.error('Erro ao buscar contagem:', err);
    res.status(500).json({ error: 'Erro ao buscar contagem' });
  }
});

app.listen(port, () => {
  console.log(`Servidor rodando em http://localhost:${port}`);
});
  
// Rota para exportar para Excel com filtros aplicados
app.get('/api/equipamentos/export', async (req, res) => {
    try {
      const { dataInicial, dataFinal, equipamento } = req.query;
      
      const pool = await sql.connect(config);
      const request = pool.request();

      let query = `
        SELECT 
            [Instalação],
            [Nota],
            [Cliente],
            [Texto breve para o code],
            [Alavanca],
            CONVERT(VARCHAR, [Data Conclusão], 120) AS [Data Conclusão],
            [Equipamento Removido],
            [Material Removido],
            [Descrição Mat. Removido],
            [Status Equip. Removido],
            [Equipamento Instalado],
            [Material Instalado],
            [Descrição Mat. Instalado],
            [Status Equip. Instalado]
        FROM [dbo].[vw_equipe_removido]
        WHERE 1=1
      `;
  
      // Aplica filtro de datas
      if (dataInicial && dataFinal) {
        query += ` AND [Data Conclusão] BETWEEN @dataInicial AND @dataFinal`;
        request.input('dataInicial', sql.Date, dataInicial);
        request.input('dataFinal', sql.Date, dataFinal);
      }
  
      // Aplica filtro de equipamentos (com tratamento de segurança)
      if (equipamento) {
        const equipamentos = equipamento.split(',')
          .map(e => e.trim())
          .filter(e => e !== '');
  
        if (equipamentos.length > 0) {
          if (equipamentos.length === 1) {
            query += ` AND [Equipamento Removido] = @equipamento`;
            request.input('equipamento', sql.NVarChar, equipamentos[0]);
          } else {
            // Cria lista de parâmetros dinâmicos para múltiplos valores
            const paramsList = equipamentos.map((e, i) => {
              const paramName = `equip${i}`;
              request.input(paramName, sql.NVarChar, e);
              return `@${paramName}`;
            }).join(',');
            
            query += ` AND [Equipamento Removido] IN (${paramsList})`;
          }
        }
      }
  
      // Ordenação
      query += ` ORDER BY [Data Conclusão] DESC`;
  
      // Log para depuração
      console.log('Executando query:', query);
      console.log('Parâmetros:', {
        dataInicial,
        dataFinal,
        equipamento
      });
  
      // Executa a query
      const result = await request.query(query);
  
      // Criação do arquivo Excel
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Sistema de Consulta';
      workbook.created = new Date();
      
      const worksheet = workbook.addWorksheet('Equipamentos');
      
      // Configuração das colunas
      worksheet.columns = [
        { header: 'Instalação', key: 'Instalação', width: 15 },
        { header: 'Nota', key: 'Nota', width: 12 },
        { header: 'Cliente', key: 'Cliente', width: 30 },
        { header: 'Descrição Nota', key: 'Texto breve para o code', width: 40 },
        { header: 'Alavanca', key: 'Alavanca', width: 15 },
        { 
          header: 'Data Conclusão', 
          key: 'Data Conclusão', 
          width: 15, 
          style: { numFmt: 'dd/mm/yyyy' } 
        },
        { header: 'Equipamento Removido', key: 'Equipamento Removido', width: 20 },
        { header: 'Material Removido', key: 'Material Removido', width: 20 },
        { header: 'Descrição Mat. Removido', key: 'Descrição Mat. Removido', width: 20 },
        { header: 'Status Equip. Removido', key: 'Status Equip. Removido', width: 20 },
        { header: 'Equipamento Instalado', key: 'Equipamento Instalado', width: 20 },
        { header: 'Material Instalado', key: 'Material Instalado', width: 20 },
        { header: 'Descrição Mat. Instalado', key: 'Descrição Mat. Instalado', width: 20 },
        { header: 'Status Equip. Instalado', key: 'Status Equip. Instalado', width: 20 },
      ];
      
      // Adiciona os dados
      if (result.recordset && result.recordset.length > 0) {
        worksheet.addRows(result.recordset);
      } else {
        worksheet.addRow(['Nenhum dado encontrado com os filtros aplicados']);
      }
  
      // Formatação dos cabeçalhos
      const headerRow = worksheet.getRow(1);
      headerRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FF1976D2' } // Azul do MUI
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF000000' } },
          left: { style: 'thin', color: { argb: 'FF000000' } },
          bottom: { style: 'thin', color: { argb: 'FF000000' } },
          right: { style: 'thin', color: { argb: 'FF000000' } }
        };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });
  
      // Configura a resposta
      res.setHeader('Content-Type', 
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 
        'attachment; filename=equipamentos_filtrados.xlsx');
      
      // Gera e envia o arquivo
      await workbook.xlsx.write(res);
      res.end();
      
    } catch (err) {
      console.error('Erro durante a exportação:', {
        message: err.message,
        stack: err.stack,
        query: req.query
      });
      
      res.status(500).json({ 
        error: 'Erro durante a exportação',
        details: process.env.NODE_ENV === 'development' ? err.message : 'Ocorreu um erro'
      });
    }
  });
  
  // Inicia o servidor
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => {
    console.log(`Servidor rodando na porta ${PORT}`);
  });