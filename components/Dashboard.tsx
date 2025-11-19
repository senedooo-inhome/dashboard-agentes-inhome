import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { downloadCSV } from '../utils/helpers';
import { CallRecord } from '../types';
import {
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  BarChart,
  Bar,
  LabelList,
} from 'recharts';
import { LogOut, Download, RefreshCw, Filter } from 'lucide-react';

/* Props */
interface DashboardProps {
  agentId: string;
  onLogout: () => void;
}

/* Helpers */
const normalizeStatus = (status?: string) =>
  (status || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

const isAnswered = (status?: string) => {
  const s = normalizeStatus(status);
  if (!s) return false;
  if (s.startsWith('nao atendida')) return false;
  return s.startsWith('atendida');
};

const isNotAnswered = (status?: string) => {
  const s = normalizeStatus(status);
  if (!s) return false;
  if (s.startsWith('nao atendida') || s.includes('perdida') || s.includes('abandonada'))
    return true;
  return !isAnswered(status);
};

const formatDuration = (value?: string) => {
  if (!value) return '00:00';
  const parts = value.split(':');
  if (parts.length === 3) {
    const total = Number(parts[0]) * 3600 + Number(parts[1]) * 60 + Number(parts[2]);
    const mm = String(Math.floor(total / 60)).padStart(2, '0');
    const ss = String(total % 60).padStart(2, '0');
    return `${mm}:${ss}`;
  }
  if (parts.length === 2) return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}`;
  return '00:00';
};

/* Componente */
const Dashboard: React.FC<DashboardProps> = ({ agentId, onLogout }) => {
  /* Permissão automática */
  const supervisorIds = ['517', '307'];
  const canManageFiles = supervisorIds.includes(agentId.toString());

  /* Estados */
  const [allData, setAllData] = useState<CallRecord[]>([]);
  const [data, setData] = useState<CallRecord[]>([]);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const currentYear = new Date().getFullYear();
  const currentMonth = new Date().getMonth() + 1;
  const [selectedYear, setSelectedYear] = useState(currentYear);
  const [selectedMonth, setSelectedMonth] = useState(currentMonth);
  const lastUpdated = localStorage.getItem('callData:lastUpdated');

  /* Carregar dados do localStorage */
  useEffect(() => {
    try {
      const saved = localStorage.getItem('callData');
      if (saved) {
        const parsed: CallRecord[] = JSON.parse(saved);
        setAllData(parsed);
        console.log('Dados carregados do localStorage:', parsed.length);
      }
    } catch (err) {
      console.error('Erro ao carregar storage:', err);
    }
  }, []);

  /* Upload Excel */
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target?.result;
        const workbook = XLSX.read(buffer, { type: 'array' });
        const sheet = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheet];

        const json: any[] = XLSX.utils.sheet_to_json(worksheet, { raw: false, defval: '' });

        const parsed: CallRecord[] = json.map((row) => ({
          data_hora:
            row['Data/hora'] ||
            row['Data/Hora'] ||
            row['DATA_HORA'] ||
            '',
          did: row['DID'] || '',
          ramal: row['Ramal']?.toString() || row['RAMAL']?.toString() || '',
          alias:
            row['Ramal']?.toString() ||
            row['RAMAL']?.toString() ||
            row['Agente']?.toString() ||
            '',
          agente: row['Agente']?.toString() || '',
          status: row['Status'] || '',
          fila: row['Fila'] || '',
          tabulacao: row['Tabulação'] || row['Tabulacao'] || '',
          origem: row['Origem'] || '',
          tempo_falado:
            row['Tempo falado'] ||
            row['Tempo Falado'] ||
            row['Duração'] ||
            row['Duracao'] ||
            '',
          uniqueid: row['uniqueid']?.toString() || '',
        }));

        setAllData(parsed);

        /* salvar no localStorage */
        localStorage.setItem('callData', JSON.stringify(parsed));
        localStorage.setItem('callData:lastUpdated', new Date().toISOString());
      } catch (err) {
        console.error('Erro ao ler Excel:', err);
      } finally {
        setLoading(false);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  /* Filtro por agente + mês/ano */
  useEffect(() => {
    if (!allData.length) return setData([]);

    const filtered = allData.filter((call) => {
      const match =
        call.ramal === agentId ||
        call.agente === agentId ||
        call.alias === agentId;

      if (!match) return false;

      const raw = call.data_hora?.split(' ')[0];
      if (!raw) return false;

      const [d, m, y] = raw.split('/');
      const finalYear = y.length === 2 ? Number('20' + y) : Number(y);
      const callDate = new Date(finalYear, Number(m) - 1, Number(d));

      return (
        callDate.getFullYear() === selectedYear &&
        callDate.getMonth() + 1 === selectedMonth
      );
    });

    setData(filtered);
  }, [allData, agentId, selectedMonth, selectedYear]);

  /* Estatísticas */
  const stats = useMemo(() => {
    const daily = new Map<string, any>();
    const queue = new Map<string, number>();
    let totalA = 0;
    let totalN = 0;
    let total = 0;

    data.forEach((call) => {
      total++;

      const answered = isAnswered(call.status);
      const notAnswered = isNotAnswered(call.status);

      if (answered) totalA++;
      if (notAnswered) totalN++;

      const date = call.data_hora.split(' ')[0];
      if (!daily.has(date)) daily.set(date, { total: 0, atendidas: 0, naoAtendidas: 0 });
      const d = daily.get(date);
      d.total++;
      if (answered) d.atendidas++;
      if (notAnswered) d.naoAtendidas++;

      queue.set(call.fila || 'Sem Fila', (queue.get(call.fila || 'Sem Fila') || 0) + 1);
    });

    return {
      total,
      totalA,
      totalN,
      lineChart: [...daily.entries()].map(([date, o]) => ({ date, ...o })),
      barChart: [...queue.entries()].map(([name, value]) => ({ name, value })),
    };
  }, [data]);

  /* Export CSV */
  const handleExport = () => {
    downloadCSV(data, `relatorio_${agentId}_${selectedMonth}_${selectedYear}.csv`);
  };

  return (
    <div className="min-h-screen bg-gray-100 pb-10">
      {/* Cabeçalho */}
      <header className="bg-white shadow-sm sticky top-0 z-20">
        <div className="max-w-7xl mx-auto px-4 h-16 flex justify-between items-center">
          <div className="flex items-center gap-2">
            <div className="bg-blue-600 text-white p-1.5 rounded font-bold text-lg">SX</div>
            <h1 className="text-xl font-semibold text-gray-800 hidden sm:block">
              Dashboard do Agente
            </h1>
          </div>

          <div className="flex items-center gap-4">
            {lastUpdated && (
              <span className="text-xs text-gray-500">
                Atualizado em: {new Date(lastUpdated).toLocaleString('pt-BR')}
              </span>
            )}

            <span className="text-sm bg-gray-100 px-3 py-1 rounded-full">
              Agente: <strong>{agentId}</strong>
            </span>

            <button onClick={onLogout} className="text-gray-500 hover:text-red-600 p-2">
              <LogOut />
            </button>
          </div>
        </div>
      </header>

      {/* Conteúdo */}
      <main className="max-w-7xl mx-auto px-4 mt-8">
        {/* Filtros */}
        <div className="bg-white p-4 rounded-xl shadow-sm border mb-8 flex flex-col sm:flex-row justify-between items-center gap-4">
          <div className="flex items-center gap-3">
            {/* Filtros de ano/mês */}
            <div className="flex items-center gap-2 bg-gray-50 px-3 py-2 rounded-lg border">
              <Filter className="h-4 w-4 text-gray-400" />
              <select
                value={selectedYear}
                onChange={(e) => setSelectedYear(Number(e.target.value))}
                className="bg-transparent"
              >
                <option value={currentYear}>{currentYear}</option>
                <option value={currentYear - 1}>{currentYear - 1}</option>
                <option value={currentYear - 2}>{currentYear - 2}</option>
              </select>
              <select
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(Number(e.target.value))}
                className="bg-transparent"
              >
                {Array.from({ length: 12 }).map((_, i) => (
                  <option key={i} value={i + 1}>
                    {i + 1}
                  </option>
                ))}
              </select>
            </div>

            <button onClick={() => setData([...data])} className="p-2 rounded">
              <RefreshCw className="h-5 w-5" />
            </button>

            {/* Upload – SOMENTE supervisores */}
            {canManageFiles && (
              <label className="cursor-pointer flex items-center bg-gray-100 px-3 py-2 rounded-lg border border-dashed">
                Escolher Excel
                <input type="file" className="hidden" accept=".xlsx,.xls" onChange={handleFileChange} />
              </label>
            )}
          </div>

          {/* Baixar CSV – SOMENTE supervisores */}
          {canManageFiles && (
            <button
              onClick={handleExport}
              className="bg-green-600 text-white px-4 py-2 rounded-lg shadow"
            >
              <Download className="h-4 w-4" /> Exportar CSV
            </button>
          )}
        </div>

        {/* Cards */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <div className="bg-white p-6 rounded-xl shadow border">
            <p>Total de ligações</p>
            <h2 className="text-3xl font-bold">{stats.total}</h2>
          </div>
          <div className="bg-white p-6 rounded-xl shadow border">
            <p>Atendidas</p>
            <h2 className="text-3xl font-bold text-green-600">{stats.totalA}</h2>
          </div>
          <div className="bg-white p-6 rounded-xl shadow border">
            <p>Não atendidas</p>
            <h2 className="text-3xl font-bold text-red-600">{stats.totalN}</h2>
          </div>
        </div>

        {/* Gráfico Linha */}
        <div className="bg-white p-6 rounded-xl shadow border mb-8">
          <h3 className="font-semibold mb-4">Evolução Diária</h3>
          <ResponsiveContainer width="100%" height={300}>
            <LineChart data={stats.lineChart}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="date" />
              <YAxis />
              <Tooltip />
              <Legend />
              <Line type="monotone" dataKey="atendidas" stroke="#10b981" />
              <Line type="monotone" dataKey="naoAtendidas" stroke="#ef4444" />
            </LineChart>
          </ResponsiveContainer>
        </div>

        {/* Gráfico Barras */}
        <div className="bg-white p-6 rounded-xl shadow border mb-8">
          <h3 className="font-semibold mb-4">Volume por Fila</h3>
          <ResponsiveContainer width="100%" height={350}>
            <BarChart data={stats.barChart} layout="vertical">
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis type="number" />
              <YAxis type="category" dataKey="name" width={200} />
              <Bar dataKey="value" fill="#3b82f6">
                <LabelList dataKey="value" position="right" />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        {/* Tabela */}
        <div className="bg-white rounded-xl shadow border overflow-hidden mb-10">
          <table className="min-w-full text-sm">
            <thead className="bg-gray-100">
              <tr>
                <th className="px-4 py-2">Data</th>
                <th className="px-4 py-2">Fila</th>
                <th className="px-4 py-2">Status</th>
                <th className="px-4 py-2">Duração</th>
              </tr>
            </thead>

            <tbody>
              {data.map((call, i) => {
                const answered = isAnswered(call.status);
                const notAnswered = isNotAnswered(call.status);

                return (
                  <tr key={i} className="border-t hover:bg-gray-50">
                    <td className="px-4 py-2">{call.data_hora}</td>
                    <td className="px-4 py-2">{call.fila}</td>
                    <td className="px-4 py-2">
                      <span
                        className={`px-2 py-1 rounded text-xs ${
                          answered
                            ? 'bg-green-100 text-green-700'
                            : notAnswered
                            ? 'bg-red-100 text-red-700'
                            : 'bg-gray-100 text-gray-800'
                        }`}
                      >
                        {call.status}
                      </span>
                    </td>
                    <td className="px-4 py-2">{formatDuration(call.tempo_falado)}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </main>
    </div>
  );
};

export default Dashboard;
