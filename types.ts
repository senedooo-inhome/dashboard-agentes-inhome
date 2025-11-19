export interface LoginResponse {
  status: number;
  token: string;
  msg?: string;
}

export interface CallRecord {
  data_hora: string;   // Data/hora
  did: string;         // DID
  ramal: string;       // Ramal (identifica o agente)
  alias: string;       // Alias usado p/ filtrar pelo agente

  agente: string;      // Agente (nome ou código)
  status: string;      // Status da chamada
  fila: string;        // Fila de atendimento
  tabulacao: string;   // Tabulação final da chamada

  // 👇 esses aqui são usados no Dashboard.tsx
  origem?: string;        // Número de origem / callerid
  tempo_falado?: string;  // Duração da chamada
  uniqueid?: string;      // ID único (se vier do sistema)
}


export interface FilterState {
  year: number;
  month: number;
}

export interface ChartData {
  name: string;
  value: number;
  color?: string;
}

export interface DailyStat {
  date: string;
  total: number;
  atendidas: number;
  naoAtendidas: number;
}