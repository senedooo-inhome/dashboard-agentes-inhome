import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { getMonthDateRange, downloadCSV } from '../utils/helpers';
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

interface DashboardProps {
  agentId: string;
  onLogout: () => void;
}

const COLORS = ['#3b82f6', '#ef4444', '#10b981', '#f59e0b'];

// --------- Helpers de status / duração ----------
const normalizeStatus = (status?: string) =>
  (status || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

const isAnswered = (status?: string) => {
  const s = normalizeStatus(status);
  if (!s) return false;

  // importante: "não atendida" contém "atendida"
  if (s.startsWith('nao atendida')) return false;

  return s.startsWith('atendida');
};

const isNotAnswered = (status?: string) => {
  const s = normalizeStatus(status);
  if (!s) return false;

  if (
    s.startsWith('nao atendida') ||
    s.includes('perdida') ||
    s.includes('abandonada')
  ) {
    return true;
  }

  return !isAnswered(status);
};

const formatDuration = (value?: string) => {
  if (!value) return '00:00';
  const v = value.trim();
  if (!v) return '00:00';

  // já vem em HH:MM:SS ou MM:SS
  if (v.includes(':')) {
    const parts = v.split(':').map((p) => p.padStart(2, '0'));

    if (parts.length === 3) {
      const [h, m, s] = parts;
      const totalSeconds = Number(h) * 3600 + Number(m) * 60 + Number(s);
      const mm = Math.floor(totalSeconds / 60)
        .toString()
        .padStart(2, '0');
      const ss = (totalSeconds % 60).toString().padStart(2, '0');
      return `${mm}:${ss}`;
    }

    if (parts.length === 2) {
      const [m, s] = parts;
      return `${m.padStart(2, '0')}:${s.padStart(2, '0')}`;
    }

    return v;
  }

  // só em segundos
  const seconds = Number(v.replace(',', '.'));
  if (!Number.isFinite(seconds)) return v;

  const mm = Math.floor(seconds / 60)
    .toString()
    .padStart(2, '0');
  const ss = (seconds % 60).toString().padStart(2, '0');
  return `${mm}:${ss}`;
};

const Dashboard: React.FC<DashboardProps> = ({ agentId, onLogout }) => {
  // supervisores que podem gerenciar arquivos (upload + CSV)
  const supervisorIds = ['517', '307'];
  const canManageFiles = supervisorIds.includes(agentId.toString());

  const [allData, setAllData] = useState<CallRecord[]>([]);
  const [data, setData] = useState<CallRecord[]>([]);
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);

  const currentYear = new Date().getFullYear();
  const currentMonth = new Date().getMonth() + 1;

  const [selectedYear, setSelectedYear] = useState(currentYear);
  const [selectedMonth, setSelectedMonth] = useState(currentMonth);

  const months = [
    { val: 1, name: 'Janeiro' },
    { val: 2, name: 'Fevereiro' },
    { val: 3, name: 'Março' },
    { val: 4, name: 'Abril' },
    { val: 5, name: 'Maio' },
    { val: 6, name: 'Junho' },
    { val: 7, name: 'Julho' },
    { val: 8, name: 'Agosto' },
    { val: 9, name: 'Setembro' },
    { val: 10, name: 'Outubro' },
    { val: 11, name: 'Novembro' },
    { val: 12, name: 'Dezembro' },
  ];

  const years = Array.from({ length: 5 }, (_, i) => currentYear - i);

  // ---------- Restauração automática do localStorage ----------
  useEffect(() => {
    try {
      const saved = localStorage.getItem('callData');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          console.log(
            '✔️ Dados carregados do localStorage:',
            parsed.length,
            'registros'
          );
          setAllData(parsed);
          setFileName('Carregado automaticamente (localStorage)');
        }
      }
    } catch (err) {
      console.error('Erro ao restaurar do localStorage:', err);
    }
  }, []);

  // --------- Leitura do Excel local ----------
  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setFileName(file.name);

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) return;

        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const json: any[] = XLSX.utils.sheet_to_json(worksheet, {
          raw: false,
          defval: '',
        });

        console.log('Primeira linha bruta do Excel:', json[0]);
        console.log('Cabeçalhos:', json[0] ? Object.keys(json[0]) : []);

        const parsed: CallRecord[] = json.map((row, index) => {
          const rec: CallRecord = {
            data_hora:
              row['Data/hora'] ||
              row['Data/Hora'] ||
              row['DATA_HORA'] ||
              '',
            did: row['DID'] || '',
            ramal:
              row['Ramal']?.toString() ||
              row['RAMAL']?.toString() ||
              '',
            alias:
              row['Ramal']?.toString() ||
              row['RAMAL']?.toString() ||
              row['Agente']?.toString() ||
              row['AGENTE']?.toString() ||
              '',
            agente:
              row['Agente']?.toString() ||
              row['AGENTE']?.toString() ||
              '',
            status: row['Status'] || row['STATUS'] || '',
            fila: row['Fila'] || row['FILA'] || '',
            tabulacao:
              row['Tabulação'] ||
              row['Tabulacao'] ||
              row['TABULACAO'] ||
              '',
            origem: row['Origem'] || row['ORIGEM'] || '',
            tempo_falado:
              row['Tempo falado'] ||
              row['Tempo Falado'] ||
              row['Duração'] ||
              row['Duracao'] ||
              row['Duração da chamada'] ||
              row['TEMPO_FALADO'] ||
              '',
            uniqueid:
              row['uniqueid']?.toString() ||
              row['UNIQUEID']?.toString() ||
              '',
          };

          if (index < 5) {
            console.log(`Linha ${index} (row):`, row);
            console.log(`Linha ${index} (parsed):`, rec);
          }

          return rec;
        });

        console.log('Total de registros lidos do Excel:', parsed.length);

        // guarda na memória da aplicação
        setAllData(parsed);

        // salva no navegador para não precisar fazer upload novamente
        try {
          localStorage.setItem('callData', JSON.stringify(parsed));
          localStorage.setItem(
            'callData:lastUpdated',
            new Date().toISOString()
          );
        } catch (err) {
          console.error('Erro ao salvar no localStorage:', err);
        }
      } catch (err) {
        console.error('Erro ao ler o arquivo Excel:', err);
      } finally {
        setLoading(false);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  // --------- Filtro por agente + mês/ano ----------
  useEffect(() => {
    if (!allData.length) {
      setData([]);
      return;
    }

    const filtered = allData.filter((call) => {
      const agentMatch =
        call.ramal?.toString() === agentId ||
        call.agente?.toString() === agentId ||
        call.alias?.toString() === agentId;

      if (!agentMatch) return false;

      const rawDateTime = call.data_hora?.trim();
      if (!rawDateTime) return false;

      const datePart = rawDateTime.split(' ')[0];
      if (!datePart || !datePart.includes('/')) return false;

      const [p1, p2, p3] = datePart.split('/');

      let day = p1;
      let month = p2;
      let year = p3;

      if (Number(p1) <= 12 && Number(p2) > 12) {
        month = p1;
        day = p2;
      }

      if (year.length === 2) {
        year = '20' + year;
      }

      const callDate = new Date(
        Number(year),
        Number(month) - 1,
        Number(day)
      );

      const monthMatch = callDate.getMonth() + 1 === selectedMonth;
      const yearMatch = callDate.getFullYear() === selectedYear;

      return monthMatch && yearMatch;
    });

    setData(filtered);
  }, [allData, agentId, selectedYear, selectedMonth]);

  // --------- Estatísticas / Gráficos ----------
  const stats = useMemo(() => {
    const dailyMap = new Map<
      string,
      { total: number; atendidas: number; naoAtendidas: number }
    >();
    const queueMap = new Map<string, number>();

    let totalCalls = 0;
    let totalAtendidas = 0;
    let totalNaoAtendidas = 0;

    data.forEach((call) => {
      totalCalls++;

      const answered = isAnswered(call.status);
      const notAnswered = isNotAnswered(call.status);

      if (answered) totalAtendidas++;
      else if (notAnswered) totalNaoAtendidas++;

      const dateKey = call.data_hora?.split(' ')[0] || 'Desconhecido';

      if (!dailyMap.has(dateKey)) {
        dailyMap.set(dateKey, {
          total: 0,
          atendidas: 0,
          naoAtendidas: 0,
        });
      }

      const dayStat = dailyMap.get(dateKey)!;
      dayStat.total++;

      if (answered) dayStat.atendidas++;
      else if (notAnswered) dayStat.naoAtendidas++;

      const queue = call.fila || 'Sem Fila';
      queueMap.set(queue, (queueMap.get(queue) || 0) + 1);
    });

    const lineChartData = Array.from(dailyMap.entries())
      .map(([date, stat]) => ({
        date,
        ...stat,
      }))
      .sort((a, b) => {
        const [d1, m1, y1] = a.date.split('/');
        const [d2, m2, y2] = b.date.split('/');

        return (
          new Date(`${y1}-${m1}-${d1}`).getTime() -
          new Date(`${y2}-${m2}-${d2}`).getTime()
        );
      });

    const pieChartData = [
      { name: 'Atendidas', value: totalAtendidas },
      { name: 'Não Atendidas', value: totalNaoAtendidas },
    ];

    const barChartData = Array.from(queueMap.entries()).map(
      ([name, value]) => ({
        name,
        value,
      })
    );

    return {
      totalCalls,
      totalAtendidas,
      totalNaoAtendidas,
      lineChartData,
      pieChartData,
      barChartData,
    };
  }, [data]);

  const handleExport = () => {
    downloadCSV(
      data,
      `relatorio_agente_${agentId}_${selectedMonth}_${selectedYear}.csv`
    );
  };

  const handleRefresh = () => {
    if (allData.length) {
      setData((prev) => [...prev]);
    }
  };

  return (
    <div className="min-h-screen bg-gray-100 pb-10">
      {/* Header */}
      <header className="bg-white shadow-sm sticky top-0 z-20">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex justify-between items-center">
          <div className="flex items-center gap-2">
            <div className="bg-blue-600 text-white p-1.5 rounded font-bold text-lg">
              SX
            </div>
            <h1 className="text-xl font-semibold text-gray-800 hidden sm:block">
              Dashboard do Agente
            </h1>
          </div>

          <div className="flex items-center gap-4">
            <span className="text-sm text-gray-600 bg-gray-100 px-3 py-1 rounded-full">
              Agente:{' '}
              <span className="font-medium text-gray-900">
                {agentId}
              </span>
            </span>
            <button
              onClick={onLogout}
              className="text-gray-500 hover:text-red-600 transition-colors p-2 rounded-full hover:bg-red-50"
              title="Sair"
            >
              <LogOut className="h-5 w-5" />
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 mt-8">
        {/* Controls */}
        <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-200 mb-8 flex flex-col sm:flex-row justify-between items-center gap-4">
          <div className="flex flex-col sm:flex-row items-center gap-3 w-full sm:w-auto">
            <div className="flex items-center gap-2 bg-gray-50 px-3 py-2 rounded-lg border border-gray-200">
              <Filter className="h-4 w-4 text-gray-400" />
              <select
                value={selectedYear}
                onChange={(e) =>
                  setSelectedYear(Number(e.target.value))
                }
                className="bg-transparent border-none text-sm font-medium focus:ring-0 text-gray-700"
              >
                {years.map((y) => (
                  <option key={y} value={y}>
                    {y}
                  </option>
                ))}
              </select>
              <span className="text-gray-300">|</span>
              <select
                value={selectedMonth}
                onChange={(e) =>
                  setSelectedMonth(Number(e.target.value))
                }
                className="bg-transparent border-none text-sm font-medium focus:ring-0 text-gray-700"
              >
                {months.map((m) => (
                  <option key={m.val} value={m.val}>
                    {m.name}
                  </option>
                ))}
              </select>
            </div>

            <button
              onClick={handleRefresh}
              className="p-2 text-blue-600 hover:bg-blue-50 rounded-lg transition-colors"
              title="Reaplicar filtros"
            >
              <RefreshCw
                className={`h-5 w-5 ${
                  loading ? 'animate-spin' : ''
                }`}
              />
            </button>

            {/* Upload do Excel – só pra quem pode gerenciar arquivos */}
            {canManageFiles && (
              <label className="flex items-center gap-2 text-sm text-gray-700 cursor-pointer">
                <span className="bg-gray-100 px-3 py-2 rounded-lg border border-dashed border-gray-300 hover:bg-gray-50">
                  Escolher arquivo Excel
                </span>
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileChange}
                  className="hidden"
                />
                {fileName && (
                  <span className="text-xs text-gray-500 truncate max-w-[150px]">
                    {fileName}
                  </span>
                )}
              </label>
            )}
          </div>

          {/* Botão Baixar CSV – só pra quem pode também */}
          {canManageFiles && (
            <button
              onClick={handleExport}
              disabled={data.length === 0}
              className="flex items-center gap-2 bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed text-sm font-medium w-full sm:w-auto justify-center"
            >
              <Download className="h-4 w-4" />
              Baixar CSV
            </button>
          )}
        </div>

        {/* Summary Cards */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
            <p className="text-sm font-medium text-gray-500">
              Total de Ligações
            </p>
            <p className="text-3xl font-bold text-gray-900 mt-2">
              {stats.totalCalls}
            </p>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-blue-100 bg-blue-50">
            <p className="text-sm font-medium text-blue-600">
              Atendidas
            </p>
            <p className="text-3xl font-bold text-blue-700 mt-2">
              {stats.totalAtendidas}
            </p>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-red-100 bg-red-50">
            <p className="text-sm font-medium text-red-600">
              Não Atendidas
            </p>
            <p className="text-3xl font-bold text-red-700 mt-2">
              {stats.totalNaoAtendidas}
            </p>
          </div>
        </div>

        {/* Charts Grid */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
          {/* Line Chart */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200 lg:col-span-2">
            <h3 className="text-lg font-semibold text-gray-800 mb-6">
              Evolução Diária
            </h3>
            <ResponsiveContainer width="100%" height={300}>
              <LineChart data={stats.lineChartData}>
                <CartesianGrid
                  strokeDasharray="3 3"
                  stroke="#f0f0f0"
                />
                <XAxis
                  dataKey="date"
                  stroke="#9ca3af"
                  tick={{ fontSize: 12 }}
                />
                <YAxis
                  stroke="#9ca3af"
                  tick={{ fontSize: 12 }}
                />
                <Tooltip
                  contentStyle={{
                    borderRadius: '8px',
                    border: 'none',
                    boxShadow:
                      '0 4px 6px -1px rgba(0, 0, 0, 0.1)',
                  }}
                />
                <Legend />
                <Line
                  type="monotone"
                  dataKey="atendidas"
                  name="Atendidas"
                  stroke="#10b981"
                  strokeWidth={2}
                  dot={{ r: 4 }}
                  activeDot={{ r: 6 }}
                />
                <Line
                  type="monotone"
                  dataKey="naoAtendidas"
                  name="Perdidas"
                  stroke="#ef4444"
                  strokeWidth={2}
                  dot={{ r: 4 }}
                  activeDot={{ r: 6 }}
                />
              </LineChart>
            </ResponsiveContainer>
          </div>

          {/* Bar Chart */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
            <h3 className="text-lg font-semibold text-gray-800 mb-6">
              Volume por Fila
            </h3>
            <div className="h-80">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={stats.barChartData}
                  layout="vertical"
                  margin={{
                    top: 10,
                    right: 30,
                    left: 200,
                    bottom: 10,
                  }}
                >
                  <CartesianGrid
                    strokeDasharray="3 3"
                    stroke="#f0f0f0"
                  />
                  <XAxis
                    type="number"
                    stroke="#9ca3af"
                    tick={{ fontSize: 12 }}
                  />
                  <YAxis
                    type="category"
                    dataKey="name"
                    stroke="#9ca3af"
                    tick={{ fontSize: 12 }}
                    width={220}
                  />
                  <Tooltip cursor={{ fill: '#f9fafb' }} />
                  <Bar
                    dataKey="value"
                    name="Ligações"
                    fill="#3b82f6"
                    radius={[0, 4, 4, 0]}
                  >
                    <LabelList
                      dataKey="value"
                      position="right"
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {/* Detailed Table */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden mb-10">
          <div className="px-6 py-4 border-b border-gray-200">
            <h3 className="text-lg font-semibold text-gray-800">
              Detalhamento de Chamadas
            </h3>
          </div>
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200">
              <thead className="bg-gray-50">
                <tr>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Data/Hora
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Fila
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Status
                  </th>
                  <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Duração
                  </th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {data.length > 0 ? (
                  data.map((row, idx) => {
                    const answered = isAnswered(row.status);
                    const notAnswered = isNotAnswered(row.status);

                    return (
                      <tr
                        key={row.uniqueid || idx}
                        className="hover:bg-gray-50 transition-colors"
                      >
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                          {row.data_hora ||
                            'Data não disponível'}
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-600">
                          {row.fila || 'Sem Fila'}
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <span
                            className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${
                              answered
                                ? 'bg-green-100 text-green-800'
                                : notAnswered
                                ? 'bg-red-100 text-red-800'
                                : 'bg-gray-100 text-gray-800'
                            }`}
                          >
                            {row.status || 'Sem status'}
                          </span>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm text-gray-600">
                          {formatDuration(row.tempo_falado)}
                        </td>
                      </tr>
                    );
                  })
                ) : (
                  <tr>
                    <td
                      colSpan={4}
                      className="px-6 py-8 text-center text-sm text-gray-500"
                    >
                      {loading
                        ? 'Carregando arquivo...'
                        : 'Nenhum dado encontrado. Selecione um arquivo Excel para começar.'}
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      </main>
    </div>
  );
};

export default Dashboard;
