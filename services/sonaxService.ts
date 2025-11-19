import { CallRecord } from '../types';
import users from '../src/data/users.json'; // ajuste o caminho conforme a localização do arquivo

const BASE_URL = 'https://apiv2.sonax.net.br';

// Login local baseado em users.json
export const loginAgent = async (id: string, senha: string): Promise<string> => {
  const user = users.find(u => u.id === id && u.senha === senha);
  if (!user) {
    throw new Error('Credenciais inválidas');
  }

  // Retorna um token fictício baseado no ID
  return `token-${id}`;
};

// Busca os dados reais da API
export const fetchReport = async (token: string, dt_inicio: string, dt_fim: string): Promise<CallRecord[]> => {
  const params = new URLSearchParams();
  params.append('dt_inicio', dt_inicio);
  params.append('dt_fim', dt_fim);

  try {
    const response = await fetch(`${BASE_URL}/api/vingadora/relatorioEntrante`, {
      method: 'POST',
      headers: {
        'token': token,
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: params,
    });

    if (!response.ok) {
      throw new Error(`Erro na requisição: ${response.status}`);
    }

    const jsonData = await response.json();

    if (Array.isArray(jsonData)) return jsonData;
    if (jsonData.relatorio_entrante) return jsonData.relatorio_entrante;
    if (jsonData.data) return jsonData.data;

    return [];
  } catch (error) {
    console.error("Fetch report error:", error);
    return [];
  }
};