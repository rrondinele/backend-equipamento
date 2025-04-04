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
import { ptBR } from 'date-fns/locale'; // Opcional - para formato em português



function App() {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [startDate, setStartDate] = useState(null);
  const [endDate, setEndDate] = useState(null);
  const [equipamentoFilter, setEquipamentoFilter] = useState('');

const fetchData = async () => {
  try {
    setLoading(true);
    const params = {};
    
    // Processa a lista de equipamentos (um por linha ou separados por espaço/vírgula)
    if (equipamentoFilter) {
      const equipamentos = equipamentoFilter
        .split(/[\n,\s]+/) // Divide por quebra de linha, vírgula ou espaço
        .filter(e => e.trim() !== ''); // Remove valores vazios
      
      params.equipamento = equipamentos.join(',');
    }
    
    if (startDate && endDate) {
      params.dataInicial = format(startDate, 'yyyy-MM-dd');
      params.dataFinal = format(endDate, 'yyyy-MM-dd');
    }

    const response = await axios.get('http://localhost:5000/api/equipamentos', { params });
    setData(response.data);
  } catch (err) {
    setError(err.message);
  } finally {
    setLoading(false);
  }
};

const handleExport = async () => {
  try {
    setLoading(true);
    const params = {};
    
    // Use EXATAMENTE a mesma lógica de fetchData
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

    const response = await axios.get('http://localhost:5000/api/equipamentos/export', {
      params,
      responseType: 'blob'
    });
    
    saveAs(new Blob([response.data]), 'equipamentos_filtrados.xlsx');
    setLoading(false);
  } catch (err) {
    setError(err.message);
    setLoading(false);
  }
};

  useEffect(() => {
    fetchData();
  }, []);

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
            onChange={(newValue) => setStartDate(newValue)}
            //inputFormat="dd/MM/yyyy" // Formato de exibição brasileiro
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
            onChange={(newValue) => setEndDate(newValue)}
            //inputFormat="dd/MM/yyyy" // Formato de exibição brasileiro
            renderInput={(params) => (
              <TextField 
                {...params} 
                size="small" 
                sx={{ width: 150 }}
              />
            )}
          />

          {/*
          <TextField
            label="Equip. Removido"
            value={equipamentoFilter}
            onChange={(e) => setEquipamentoFilter(e.target.value)}
            size="small"
            sx={{ width: 200 }}
            placeholder="Ex: ABC123, XYZ456"
          />
          */}

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
          
          {/* Esta Box nova é a única alteração necessária */}
          <Box sx={{ flexGrow: 1, display: 'flex', justifyContent: 'flex-end', gap: 1 }}>
            <Button 
              variant="contained" 
              onClick={fetchData}
              size="small"
              sx={{ height: 40 }}
            >
              Filtrar
            </Button>
            <Button 
              variant="outlined" 
              onClick={handleExport}
              size="small"
              sx={{ height: 40 }}
            >
              Exportar
            </Button>
          </Box>
        </Box>
      </LocalizationProvider>
  
      {/* Todo o resto permanece EXATAMENTE igual */}
      {error && (
        <Alert severity="error" sx={{ mb: 3 }}>
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
                  <TableCell>
                  {row['Data Conclusão'] ? (() => {
                    const dateStr = row['Data Conclusão'];
                    const [year, month, day] = dateStr.split('T')[0].split('-');
                    return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
                  })() : '-'}
                  </TableCell>
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