const express = require('express');
const sql = require('mssql');
const cors = require('cors');
const ExcelJS = require('exceljs');
require('dotenv').config();

// Cria a aplicação Express
const app = express();

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

// Rota para consulta específica de equipamentos
app.get('/api/equipamentos', async (req, res) => {
    try {
      const { dataInicial, dataFinal, equipamento } = req.query;
      
      const pool = await sql.connect(config);
      
      let query = `
            SELECT TOP 20 
                 [Instalação]
                ,[Nota]
                ,[Cliente]
                ,[Texto breve para o code]
                ,[Alavanca]
                ,[Data Conclusão]
                ,[Equipamento Removido]
                ,[Equipamento Instalado]
            FROM [dbo].[vw_equipe_removido]
            WHERE 1=1
      `;
  
      // Filtro por datas
      if (dataInicial && dataFinal) {
        query += ` AND [Data Conclusão] BETWEEN @dataInicial AND @dataFinal`;
      }
  
    // Filtro por equipamento (suporta múltiplos valores separados por vírgula)
    if (equipamento) {
        const equipamentos = equipamento.split(',').map(e => e.trim());
        if (equipamentos.length === 1) {
          query += ` AND [Equipamento Removido] = @equipamento`;
        } else {
          query += ` AND [Equipamento Removido] IN (${equipamentos.map(e => `'${e}'`).join(',')})`;
        }
      }
  
      query += ` ORDER BY [Data Conclusão] DESC`;
  
      const request = pool.request();
      if (dataInicial && dataFinal) {
        request.input('dataInicial', sql.Date, dataInicial);
        request.input('dataFinal', sql.Date, dataFinal);
      }
      if (equipamento && equipamento.split(',').length === 1) {
        request.input('equipamento', sql.NVarChar, equipamento.trim());
      }
  
      const result = await request.query(query);
      res.json(result.recordset);
      
    } catch (err) {
      console.error('Erro na consulta:', err);
      res.status(500).json({ 
        error: 'Erro ao consultar equipamentos',
        details: err.message 
      });
    }
  });
  
// Rota para exportar para Excel (todos os registros)
// Rota para exportar para Excel com filtros aplicados
app.get('/api/equipamentos/export', async (req, res) => {
    try {
      const { dataInicial, dataFinal, equipamento } = req.query;
      
      // Conexão com o banco de dados
      const pool = await sql.connect(config);
      const request = pool.request(); // Inicializa o request aqui
  
      // Query base
      let query = `
        SELECT 
          [Instalação],
          [Nota],
          [Cliente],
          [Texto breve para o code],
          [Alavanca],
          [Data Conclusão],
          [Equipamento Removido],
          [Equipamento Instalado]
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
        { header: 'Equipamento Instalado', key: 'Equipamento Instalado', width: 20 }
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