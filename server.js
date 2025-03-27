// App.js - Correções
import React, { useState, useEffect } from 'react';
import axios from 'axios';
import {
  Container,
  Typography,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  TextField,
  Button,
  Box,
  CircularProgress,
  Alert,
  Grid
} from '@mui/material';
import { DatePicker } from '@mui/x-date-pickers/DatePicker';
import { LocalizationProvider } from '@mui/x-date-pickers/LocalizationProvider';
import { AdapterDateFns } from '@mui/x-date-pickers/AdapterDateFns';
import { saveAs } from 'file-saver';
import { format, parseISO } from 'date-fns';
import { ptBR } from 'date-fns/locale';

// Configuração única da URL da API
const API_URL = 'https://backend-equipamento.onrender.com';

function App() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [startDate, setStartDate] = useState(null);
  const [endDate, setEndDate] = useState(null);
  const [equipamentoFilter, setEquipamentoFilter] = useState('');

  const fetchData = async () => {
    const params = {};
    
    try {
      setLoading(true);
      setError(null);
      
      if (equipamentoFilter) {
        const equipamentos = equipamentoFilter
          .split(/[\n,\s]+/)
          .filter(e => e.trim() !== '');
        
        params.equipamento = equipamentos.join(',');
      }
      
      if (startDate && endDate) {
        params.dataInicial = format(startDate, 'yyyy-MM-dd');
        params.dataFinal = format(endDate, 'yyyy-MM-dd');
      }

      const response = await axios.get(`${API_URL}/api/equipamentos`, {
        params,
        headers: {
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        timeout: 10000
      });

      if (!response.data) {
        throw new Error('Resposta da API sem dados');
      }

      setData(response.data);
      
    } catch (err) {
      const errorMessage = err.response 
        ? `Erro ${err.response.status}: ${err.response.data?.message || err.message}`
        : `Erro de conexão: ${err.message}`;
      
      setError(errorMessage);
      console.error('Erro na requisição:', {
        url: `${API_URL}/api/equipamentos`,
        error: err,
        params
      });
      
    } finally {
      setLoading(false);
    }
  };

  const handleExport = async () => {
    try {
      setLoading(true);
      const params = {};
      
      if (equipamentoFilter) {
        const equipamentos = equipamentoFilter
          .split(/[\n,\s]+/)
          .filter(e => e.trim() !== '');
        params.equipamento = equipamentos.join(',');
      }
      
      if (startDate && endDate) {
        params.dataInicial = format(startDate, 'yyyy-MM-dd');
        params.dataFinal = format(endDate, 'yyyy-MM-dd');
      }

      const response = await axios.get(`${API_URL}/api/equipamentos/export`, {
        params,
        responseType: 'blob'
      });
      
      saveAs(new Blob([response.data]), 'equipamentos_filtrados.xlsx');
      
    } catch (err) {
      const errorMessage = err.response 
        ? `Erro ${err.response.status}: ${err.response.data?.message || err.message}`
        : `Erro de exportação: ${err.message}`;
      
      setError(errorMessage);
      console.error('Erro na exportação:', err);
      
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  const formatDate = (dateStr) => {
    if (!dateStr) return '-';
    try {
      const date = new Date(dateStr);
      return isNaN(date.getTime()) ? '-' : format(date, 'dd/MM/yyyy');
    } catch {
      return '-';
    }
  };

  return (
    <Container maxWidth="lg" sx={{ mt: 4, mb: 4 }}>
      <Typography variant="h4" component="h1" gutterBottom>
        Consulta de Equipamentos
      </Typography>
      
      <LocalizationProvider dateAdapter={AdapterDateFns} adapterLocale={ptBR}>
        <Box sx={{ 
          display: 'flex', 
          gap: 1,
          mb: 2,
          alignItems: 'center',
          flexWrap: 'wrap'
        }}>
          <DatePicker
            label="Data Inicial"
            value={startDate}
            onChange={setStartDate}
            renderInput={(params) => (
              <TextField 
                {...params} 
                size="small" 
                sx={{ width: 150 }}
              />
            )}
          />
          <DatePicker
            label="Data Final"
            value={endDate}
            onChange={setEndDate}
            renderInput={(params) => (
              <TextField 
                {...params} 
                size="small" 
                sx={{ width: 150 }}
              />
            )}
          />

          <TextField
            label="Equip. Removido"
            value={equipamentoFilter}
            onChange={(e) => setEquipamentoFilter(e.target.value)}
            size="small"
            sx={{ width: 200 }}
            multiline
            rows={1.5}
            placeholder="Cole a lista aqui"
          />
          
          <Box sx={{ flexGrow: 1, display: 'flex', justifyContent: 'flex-end', gap: 1 }}>
            <Button 
              variant="contained" 
              onClick={fetchData}
              size="small"
              sx={{ height: 40 }}
              disabled={loading}
            >
              {loading ? <CircularProgress size={24} /> : 'Filtrar'}
            </Button>
            <Button 
              variant="outlined" 
              onClick={handleExport}
              size="small"
              sx={{ height: 40 }}
              disabled={loading}
            >
              {loading ? <CircularProgress size={24} /> : 'Exportar'}
            </Button>
          </Box>
        </Box>
      </LocalizationProvider>
  
      {error && (
        <Alert severity="error" sx={{ mb: 3 }} onClose={() => setError(null)}>
          {error}
        </Alert>
      )}
  
      {loading ? (
        <Box display="flex" justifyContent="center" mt={4}>
          <CircularProgress />
        </Box>
      ) : (
        <TableContainer component={Paper}>
          <Table sx={{ minWidth: 650 }} aria-label="tabela de equipamentos">
            <TableHead>
              <TableRow sx={{
                backgroundColor: '#1976d2',
                '& th': {
                  color: 'white',
                  fontWeight: '500',
                  fontSize: '0.875rem',
                  padding: '4px 16px',
                  borderBottom: 'none'
                }
              }}>
                <TableCell>Instalação</TableCell>
                <TableCell>Nota</TableCell>
                <TableCell>Cliente</TableCell>
                <TableCell>Descrição Nota</TableCell>
                <TableCell>Alavanca</TableCell>
                <TableCell>Data Conclusão</TableCell>
                <TableCell>Equipamento Removido</TableCell>
                <TableCell>Equipamento Instalado</TableCell>
              </TableRow>
            </TableHead>
            <TableBody>
              {data.map((row, index) => (
                <TableRow key={index}>
                  <TableCell>{row['Instalação']}</TableCell>
                  <TableCell>{row['Nota']}</TableCell>
                  <TableCell>{row['Cliente']}</TableCell>
                  <TableCell>{row['Texto breve para o code']}</TableCell>
                  <TableCell>{row['Alavanca']}</TableCell>
                  <TableCell>{formatDate(row['Data Conclusão'])}</TableCell>
                  <TableCell>{row['Equipamento Removido']}</TableCell>
                  <TableCell>{row['Equipamento Instalado']}</TableCell>
                </TableRow>
              ))}
            </TableBody>
          </Table>
        </TableContainer>
      )}
    </Container>
  );
}

export default App;