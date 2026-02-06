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

// 👉 TROCAR POR SEU LINK REAL DO EXCEL PUBLICADO
const EXCEL_URL =
  'https://docs.google.com/spreadsheets/d/12AnC2ureIEWxLoKbH6Yb2bSvyECDE39jIH6e_f9eJ2E/export?format=xlsx';

// ======================= LINKS GERAIS DO SISTEMA =======================
// 👉 AQUI VOCÊ ALTERA / ADICIONA LINKS GERAIS VISÍVEIS PARA TODOS OS AGENTES
const LINK_PLAYBOOK =
  'https://sites.google.com/view/playbook-in-home?usp=sharing';

const LINK_CAMPANHAS = 'https://agentes-sonax.vercel.app/campanhas';

// 👉 Link fixo da escala para o PRÓXIMO FERIADO (todos veem o mesmo)
const LINK_ESCALA_FERIADO = 'https://agentes-sonax.vercel.app/escala-feriado';

// ======================================================================

// =========== LINKS ESPECÍFICOS POR RAMAL (MONITORIA / ESCALA) ===========
// 👉 ADICIONE NOVOS RAMAIS AQUI QUANDO TIVER LINKS INDIVIDUAIS

// Monitoria de Qualidade por ramal
const MONITORIA_LINKS: Record<string, string> = {
  '523': 'https://sonax.bitrix24.com.br/~iia7C',
  '512': 'https://sonax.bitrix24.com.br/~64viI',
  '525': 'https://sonax.bitrix24.com.br/~Wkic0',
  '502': 'https://sonax.bitrix24.com.br/~76s6h',
  '524': 'https://sonax.bitrix24.com.br/~ou74E',
  '504': 'https://sonax.bitrix24.com.br/~li7CN',
  '312': 'https://sonax.bitrix24.com.br/~DI2tl',
  '366': 'https://sonax.bitrix24.com.br/~W7dnd',
  '314': 'https://sonax.bitrix24.com.br/~qkqyK',
  '511': 'https://sonax.bitrix24.com.br/~8r1oZ',
  '521': 'https://sonax.bitrix24.com.br/~UYNS1',
  '516': 'https://sonax.bitrix24.com.br/~ehWDC',
  '313': 'https://sonax.bitrix24.com.br/~XVUX3',
  '519': 'https://sonax.bitrix24.com.br/~AlFJ5',
  '359': 'https://sonax.bitrix24.com.br/~UtA0s',
  '365': 'https://sonax.bitrix24.com.br/~5qxYy',
  '323': 'https://sonax.bitrix24.com.br/~GVNPj',
  '503': 'https://sonax.bitrix24.com.br/~OoSmr',
  '515': 'https://sonax.bitrix24.com.br/~5WhSi',
  '316': 'https://sonax.bitrix24.com.br/~p9HFB',
  '509': 'https://sonax.bitrix24.com.br/~wCgBx',
  '300': 'https://sonax.bitrix24.com.br/~7CFxg',
  '518': 'https://sonax.bitrix24.com.br/~o0U2l',
  '505': 'https://sonax.bitrix24.com.br/~o2SqE',
  '361': 'https://sonax.bitrix24.com.br/~tLNaD',
  '319': 'https://sonax.bitrix24.com.br/~R0sg1',
  '305': 'https://sonax.bitrix24.com.br/~g6D9g',
  '369': 'https://sonax.bitrix24.com.br/~mwHfB',
  '362': 'https://sonax.bitrix24.com.br/~FkGoN',
  '310': 'https://sonax.bitrix24.com.br/~IyzQ2',
  '308': 'https://sonax.bitrix24.com.br/~hXc1b',
  '311': 'https://sonax.bitrix24.com.br/~u84mk',
  '500': 'https://sonax.bitrix24.com.br/~iuusg',
  '303': 'https://sonax.bitrix24.com.br/~Pf2q7',
  '320': 'https://sonax.bitrix24.com.br/~lfdLX',
  '360': 'https://sonax.bitrix24.com.br/~e1TJ6',
  '368': 'https://sonax.bitrix24.com.br/~iRjek',
  '363': 'https://sonax.bitrix24.com.br/~4s6Yf',
  '520': 'https://sonax.bitrix24.com.br/~Y38HE',
  '322': 'https://sonax.bitrix24.com.br/~Ot0El',
  '329': 'https://sonax.bitrix24.com.br/~xvyZb',
  '315': 'https://sonax.bitrix24.com.br/~qNIno',
  '514': 'https://sonax.bitrix24.com.br/~ZWbWb',
  '328': 'https://sonax.bitrix24.com.br/~hXVMP',
  '302': 'https://sonax.bitrix24.com.br/~dc7Xl',
  '325': 'https://sonax.bitrix24.com.br/~dpBqR',
  '306': 'https://sonax.bitrix24.com.br/~S0blz',
  '513': 'https://sonax.bitrix24.com.br/~LHFVG',
  '327': 'https://sonax.bitrix24.com.br/~6zw2b',
  '309': 'https://sonax.bitrix24.com.br/~8zRti',
  '364': 'https://sonax.bitrix24.com.br/~igJcN',
  '318': 'https://sonax.bitrix24.com.br/~4WR2r',
  '370': 'https://sonax.bitrix24.com.br/~ZeCBF',
  '371': 'https://sonax.bitrix24.com.br/~DtYeT',
  '372': 'https://sonax.bitrix24.com.br/~SfGqb',
  '374': 'https://sonax.bitrix24.com.br/~o2HOs',
  '367': 'https://sonax.bitrix24.com.br/~c3sEw',
  '326': 'https://sonax.bitrix24.com.br/~3wjD4',
  '377': 'https://sonax.bitrix24.com.br/~wClic',
  '376': 'https://sonax.bitrix24.com.br/~0WI6w',
  "379": 'https://sonax.bitrix24.com.br/~TRX7Y'
};

//==================================================================================================
// Escala por ramal
// Escala por ramal 6x1
const ESCALA_LINKS: Record<string, string> = {
  '302':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=704216666&single=true',
  '306':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1056206725&single=true',
  '308':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1370956757&single=true',
  '309':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1487899211&single=true',
  '315':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1663446821&single=true',
  '316':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=898449454&single=true',
  '318':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1193241344&single=true',
  '322':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=882266765&single=true',
  '323':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=55185584&single=true',
  '325':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1792621827&single=true',
  '326':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1602537472&single=true',
  '327':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=25761090&single=true',
  '328':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=552936615&single=true',
  '359':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=933802356&single=true',
  '329':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1258187118&single=true',
  '363':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=205640056&single=true',
  '364':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1191330873&single=true',
  '365':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1390493409&single=true',
  '366':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1813236963&single=true',
  '369':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=369292586&single=true',
  '371':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=82917110&single=true',
  '372':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1376758574&single=true',
  '503':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=230321399&single=true',
  '511':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=705437354&single=true',
  '514':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=593007587&single=true',
  '515':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=779322393&single=true',
  '516':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1677118619&single=true',
  '517':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=160954934&single=true',
  '519':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=642348146&single=true',
  '520':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1084940292&single=true',
  '521':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=2089755120&single=true',
  '523':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=0&single=true',
  '374':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=1261425454&single=true',
  '367':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=659513775&single=true',
  '524':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=371304553&single=true',
  '377':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=2068136846&single=true',
  '376':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=192044735&single=true',
    '379':
    'https://docs.google.com/spreadsheets/d/e/2PACX-1vS7gOJzJ0fM07hgbZm0P8vZdU8o2KkMsOsqhpmXumekTclrFrdqQSTfrdxuAEDhoIbMhSDs-xDmknxK/pubhtml?gid=383519841&single=true'
};

// ======================= NOVO: PAUSAS DURANTE A SEMANA =======================
// 517 e 523 preenchidos; demais agentes (mesmos da monitoria) ficam vazios
const PAUSAS_LINKS: Record<string, string> = {
  ...Object.fromEntries(Object.keys(MONITORIA_LINKS).map((k) => [k, ''])),

  '300': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=2142437450&single=true',
  '302': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=20507821&single=true',
  '303': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=756664457&single=true',
  '305': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=2051629859&single=true',
  '306': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=813749059&single=true',
  '308': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=102503259&single=true',
  '309': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=2136482759&single=true',
  '310': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=2023513968&single=true',
  '311': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=296449043&single=true',
  '312': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=584226349&single=true',
  '314': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=384431001&single=true',
  '315': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=820259857&single=true',
  '316': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=225579748&single=true',
  '318': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=64832934&single=true',
  '320': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1974371585&single=true',
  '322': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1542996410&single=true',
  '325': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=433543390&single=true',
  '326': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=91194220&single=true',
  '327': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1194403721&single=true',
  '328': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1363807707&single=true',

  '359': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1248875658&single=true',
  '360': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1482604308&single=true',
  '362': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1848106607&single=true',
  '363': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=383124424&single=true',
  '364': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=553908150&single=true',
  '365': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=2102730184&single=true',
  '367': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=610301336&single=true',
  '369': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1623473522&single=true',
  '370': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=244987254&single=true',
  '371': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1664629942&single=true',
  '374': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1173213530&single=true',
  '375': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=93607250&single=true',
  '376': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1695836818&single=true',
  '377': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=122679056&single=true',
  '379': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1621635781&single=true',

  '500': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=632651585&single=true',
  '502': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1239948983&single=true',
  '503': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=974466015&single=true',
  '504': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1133383432&single=true',
  '505': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1631515006&single=true',
  '509': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=421543466&single=true',
  '511': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=912023096&single=true',
  '512': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1972643601&single=true',
  '514': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1513127847&single=true',
  '515': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1009198874&single=true',
  '516': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=600735279&single=true',
  '517': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=0&single=true',
  '518': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1583816830&single=true',
  '519': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=416579432&single=true',
  '520': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1374508747&single=true',
  '521': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=2137246003&single=true',
  '523': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=151352378&single=true',
  '524': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=1949590862&single=true',
  '525': 'https://docs.google.com/spreadsheets/d/e/2PACX-1vRR3IeTTm7Lh2Xed2y0nrRp4iqtdp2tDcVWgS3nJ1vRNFNN4Y8nowZuHIDJyXWJLYD-4kCuUqkKOVxS/pubhtml?gid=797451066&single=true',
};

// =======================================================================

// ===================== FERIADOS NACIONAIS – AUTOMÁTICO =====================

type Feriado = {
  nome: string;
  data: Date;
};

const formatDatePtBR = (d: Date): string => {
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
};

const addDays = (date: Date, days: number): Date => {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
};

// Algoritmo para calcular a data da Páscoa (Meeus/Jones/Butcher)
const calcularPascoa = (year: number): Date => {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31); // 3=março, 4=abril
  const day = ((h + l - 7 * m + 114) % 31) + 1;

  return new Date(year, month - 1, day); // month-1 porque Date usa 0=jan
};

// Retorna todos os feriados (nacionais + Caratinga) de um ano
const getFeriadosNacionais = (year: number): Feriado[] => {
  const pascoa = calcularPascoa(year);
  const carnaval = addDays(pascoa, -47); // terça-feira de Carnaval
  const sextaSanta = addDays(pascoa, -2); // Paixão de Cristo
  const corpusChristi = addDays(pascoa, 60); // Corpus Christi

  const feriados: Feriado[] = [
    { nome: 'Confraternização Universal', data: new Date(year, 0, 1) }, // 01/01
    { nome: 'Carnaval', data: carnaval }, // terça-feira
    { nome: 'Paixão de Cristo', data: sextaSanta }, // sexta-feira santa
    { nome: 'Tiradentes', data: new Date(year, 3, 21) }, // 21/04
    { nome: 'Dia do Trabalho', data: new Date(year, 4, 1) }, // 01/05
    { nome: 'Corpus Christi', data: corpusChristi }, // móvel
    { nome: 'Aniversário Caratinga', data: new Date(year, 5, 24) }, // 24/06
    { nome: 'Independência do Brasil', data: new Date(year, 8, 7) }, // 07/09
    {
      nome: 'Nossa Senhora Aparecida - Padroeira do Brasil',
      data: new Date(year, 9, 12), // 12/10
    },
    { nome: 'Finados', data: new Date(year, 10, 2) }, // 02/11
    { nome: 'Proclamação da República', data: new Date(year, 10, 15) }, // 15/11
    { nome: 'Consciência Negra', data: new Date(year, 10, 20) }, // 20/11
    { nome: 'Natal', data: new Date(year, 11, 25) }, // 25/12
  ];

  return feriados.sort((a, b) => a.data.getTime() - b.data.getTime());
};

// Continua igual
const getProximoFeriado = () => {
  const hoje = new Date();
  const hojeZerado = new Date(
    hoje.getFullYear(),
    hoje.getMonth(),
    hoje.getDate()
  );

  const anoAtual = hoje.getFullYear();
  const feriadosAnoAtual = getFeriadosNacionais(anoAtual);

  const futurosAnoAtual = feriadosAnoAtual.filter(
    (f) => f.data.getTime() >= hojeZerado.getTime()
  );

  if (futurosAnoAtual.length > 0) {
    const f = futurosAnoAtual[0];
    return {
      nome: f.nome,
      data: f.data,
      dataFormatada: formatDatePtBR(f.data),
    };
  }

  // Se não tiver mais feriado no ano, pega o primeiro do próximo ano
  const feriadosProxAno = getFeriadosNacionais(anoAtual + 1);
  const primeiro = feriadosProxAno[0];
  return {
    nome: primeiro.nome,
    data: primeiro.data,
    dataFormatada: formatDatePtBR(primeiro.data),
  };
};

// ============================================================================

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

// --------- Mapeia uma linha do Excel para CallRecord ----------
const mapRowToCallRecord = (row: any, index: number): CallRecord => {
  const rec: CallRecord = {
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
      row['AGENTE']?.toString() ||
      '',
    agente: row['Agente']?.toString() || row['AGENTE']?.toString() || '',
    status: row['Status'] || row['STATUS'] || '',
    fila: row['Fila'] || row['FILA'] || '',
    tabulacao:
      row['Tabulação'] || row['Tabulacao'] || row['TABULACAO'] || '',
    origem: row['Origem'] || row['ORIGEM'] || '',
    tempo_falado:
      row['Tempo falado'] ||
      row['Tempo Falado'] ||
      row['Duração'] ||
      row['Duracao'] ||
      row['Duração da chamada'] ||
      row['TEMPO_FALADO'] ||
      '',
    uniqueid: row['uniqueid']?.toString() || row['UNIQUEID']?.toString() || '',
  };

  if (index < 5) {
    console.log(`Linha ${index} (row):`, row);
    console.log(`Linha ${index} (parsed):`, rec);
  }

  return rec;
};

const Dashboard: React.FC<DashboardProps> = ({ agentId, onLogout }) => {
  // supervisores que podem gerenciar arquivos (upload + CSV)
  const supervisorIds = ['517', '307'];
  const canManageFiles = supervisorIds.includes(agentId.toString());

  // ====== LINKS ESPECÍFICOS DO AGENTE LOGADO (MONITORIA / ESCALA / PAUSAS) ======
  const monitoriaUrl = MONITORIA_LINKS[agentId];
  const escalaUrl = ESCALA_LINKS[agentId];
  const pausasUrl = PAUSAS_LINKS[agentId]; // ✅ NOVO
  // =====================================================================

  // 👉 Próximo feriado calculado automaticamente (nacionais)
  const proximoFeriado = getProximoFeriado();

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

  // --------- Carrega dados do Excel online ----------
  const loadFromOnline = async () => {
    console.log('Tentando carregar Excel online de:', EXCEL_URL);

    const resp = await fetch(EXCEL_URL);
    if (!resp.ok) {
      throw new Error(`Erro HTTP ${resp.status}`);
    }

    const arrayBuffer = await resp.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    const json: any[] = XLSX.utils.sheet_to_json(worksheet, {
      raw: false,
      defval: '',
    });

    const parsed: CallRecord[] = json.map(mapRowToCallRecord);

    console.log('✔️ Dados carregados do Excel online:', parsed.length, 'registros');

    setAllData(parsed);
    setFileName('Carregado automaticamente (Excel online)');

    // Salva cache no localStorage
    try {
      localStorage.setItem('callData', JSON.stringify(parsed));
      localStorage.setItem('callData:lastUpdated', new Date().toISOString());
      console.log('✔️ Cache do Excel online salvo no localStorage.');
    } catch (err) {
      console.error('Erro ao salvar callData do online no localStorage:', err);
    }
  };

  // --------- Carrega dados do localStorage (fallback) ----------
  const loadFromLocalStorage = () => {
    try {
      const saved = localStorage.getItem('callData');
      if (!saved) {
        console.log('Nenhum callData encontrado no localStorage.');
        return;
      }

      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        console.log('✔️ Dados carregados do localStorage:', parsed.length, 'registros');
        setAllData(parsed);
        setFileName('Carregado automaticamente (cache localStorage)');
      } else {
        console.log('callData no localStorage está vazio ou inválido.');
      }
    } catch (err) {
      console.error('Erro ao restaurar callData do localStorage:', err);
    }
  };

  // ---------- Restaura filtros + carrega Excel ao entrar no dashboard ----------
  useEffect(() => {
    const init = async () => {
      // 1) Restaura filtros (mês/ano)
      try {
        const savedYear = localStorage.getItem('callData:selectedYear');
        const savedMonth = localStorage.getItem('callData:selectedMonth');

        if (savedYear) {
          const y = Number(savedYear);
          if (!Number.isNaN(y)) setSelectedYear(y);
        }

        if (savedMonth) {
          const m = Number(savedMonth);
          if (!Number.isNaN(m)) setSelectedMonth(m);
        }
      } catch (err) {
        console.error('Erro ao restaurar filtros do localStorage:', err);
      }

      // 2) Tenta online primeiro, se der erro usa cache
      setLoading(true);
      try {
        await loadFromOnline();
      } catch (err) {
        console.error('Erro ao carregar Excel online, tentando localStorage:', err);
        loadFromLocalStorage();
      } finally {
        setLoading(false);
      }
    };

    init();
  }, []);

  // ---------- Salva filtros no localStorage quando mudar ----------
  useEffect(() => {
    try {
      localStorage.setItem('callData:selectedYear', String(selectedYear));
      localStorage.setItem('callData:selectedMonth', String(selectedMonth));
    } catch (err) {
      console.error('Erro ao salvar filtros no localStorage:', err);
    }
  }, [selectedYear, selectedMonth]);

  // --------- Upload manual do Excel (só supervisores) ----------
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

        const parsed: CallRecord[] = json.map(mapRowToCallRecord);

        console.log('✔️ Dados carregados via upload manual:', parsed.length, 'registros');

        setAllData(parsed);

        // salva no navegador para não precisar fazer upload novamente
        try {
          localStorage.setItem('callData', JSON.stringify(parsed));
          localStorage.setItem('callData:lastUpdated', new Date().toISOString());
          console.log('✔️ Excel de upload manual salvo no localStorage.');
        } catch (err) {
          console.error('Erro ao salvar no localStorage:', err);
        }
      } catch (err) {
        console.error('Erro ao ler o arquivo Excel (upload):', err);
      } finally {
        setLoading(false);
      }
    };

    reader.readAsArrayBuffer(file);
  };

  // --------- Filtro por agente + mês/ano ----------
  useEffect(() => {
    console.log('Reaplicando filtros. Total allData:', allData.length);

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

      // trata dd/mm/yyyy vs mm/dd/yyyy
      if (Number(p1) <= 12 && Number(p2) > 12) {
        month = p1;
        day = p2;
      }

      if (year.length === 2) {
        year = '20' + year;
      }

      const callDate = new Date(Number(year), Number(month) - 1, Number(day));

      const monthMatch = callDate.getMonth() + 1 === selectedMonth;
      const yearMatch = callDate.getFullYear() === selectedYear;

      return monthMatch && yearMatch;
    });

    console.log(
      `Chamadas após filtro (agente=${agentId}, ${selectedMonth}/${selectedYear}):`,
      filtered.length
    );

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

    const barChartData = Array.from(queueMap.entries()).map(([name, value]) => ({
      name,
      value,
    }));

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
              <span className="font-medium text-gray-900">{agentId}</span>
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
        <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-200 mb-8 flex flex-col gap-4">
          <div className="flex flex-col sm:flex-row items-center gap-3 w-full sm:w-auto">
            <div className="flex items-center gap-2 bg-gray-50 px-3 py-2 rounded-lg border border-gray-200">
              <Filter className="h-4 w-4 text-gray-400" />
              <select
                value={selectedYear}
                onChange={(e) => setSelectedYear(Number(e.target.value))}
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
                onChange={(e) => setSelectedMonth(Number(e.target.value))}
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
              <RefreshCw className={`h-5 w-5 ${loading ? 'animate-spin' : ''}`} />
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

          {/* ===== ÁREA DE LINKS RÁPIDOS (PLAYBOOK / CAMPANHAS / MONITORIA / ESCALA / PAUSAS) ===== */}
          <div className="flex flex-wrap items-center gap-2">
            {/* PlayBook – link geral */}
            <a
              href={LINK_PLAYBOOK}
              target="_blank"
              rel="noopener noreferrer"
              className="text-xs sm:text-sm bg-blue-600 text-white px-3 py-2 rounded-lg hover:bg-blue-700 transition-colors"
            >
              PlayBook
            </a>

            {/* Campanhas – link geral */}
            <a
              href={LINK_CAMPANHAS}
              target="_blank"
              rel="noopener noreferrer"
              className="text-xs sm:text-sm bg-indigo-600 text-white px-3 py-2 rounded-lg hover:bg-indigo-700 transition-colors"
            >
              Campanhas
            </a>

            {/* Monitoria – aparece só se existir link para o ramal logado */}
            {monitoriaUrl && (
              <a
                href={monitoriaUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="text-xs sm:text-sm bg-green-600 text-white px-3 py-2 rounded-lg hover:bg-green-700 transition-colors"
              >
                Monitoria
              </a>
            )}

            {/* Escala – aparece só se existir link para o ramal logado */}
            {escalaUrl && (
              <a
                href={escalaUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="text-xs sm:text-sm bg-purple-600 text-white px-3 py-2 rounded-lg hover:bg-purple-700 transition-colors"
              >
                Escala
              </a>
            )}

            {/* ✅ NOVO: Suas pausas durante a semana */}
            {pausasUrl ? (
              <a
                href={pausasUrl}
                target="_blank"
                rel="noopener noreferrer"
                className="text-xs sm:text-sm bg-amber-600 text-white px-3 py-2 rounded-lg hover:bg-amber-700 transition-colors"
              >
                Horario de Pausas - SEMANA
              </a>
            ) : (
              <button
                type="button"
                disabled
                title="Link de pausas ainda não configurado para este agente."
                className="text-xs sm:text-sm bg-amber-100 text-amber-800 px-3 py-2 rounded-lg border border-amber-200 cursor-not-allowed opacity-80"
              >
                Horario de Pausas - SEMANA
              </button>
            )}

            {/* Botão Baixar CSV – só pra quem pode também */}
            {canManageFiles && (
              <button
                onClick={handleExport}
                disabled={data.length === 0}
                className="flex items-center gap-2 bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed text-xs sm:text-sm font-medium"
              >
                <Download className="h-4 w-4" />
                Baixar CSV
              </button>
            )}
          </div>
          {/* ========================================================================== */}
        </div>

        {/* Summary Cards */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 mb-8">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
            <p className="text-sm font-medium text-gray-500">Total de Ligações</p>
            <p className="text-3xl font-bold text-gray-900 mt-2">{stats.totalCalls}</p>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-blue-100 bg-blue-50">
            <p className="text-sm font-medium text-blue-600">Atendidas</p>
            <p className="text-3xl font-bold text-blue-700 mt-2">{stats.totalAtendidas}</p>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-red-100 bg-red-50">
            <p className="text-sm font-medium text-red-600">Não Atendidas</p>
            <p className="text-3xl font-bold text-red-700 mt-2">{stats.totalNaoAtendidas}</p>
          </div>
        </div>

        {/* Charts Grid */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
          {/* Line Chart - ocupa as duas colunas na primeira linha */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200 lg:col-span-2">
            <h3 className="text-lg font-semibold text-gray-800 mb-6">
              Evolução Diária
            </h3>
            <ResponsiveContainer width="100%" height={300}>
              <LineChart data={stats.lineChartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
                <XAxis dataKey="date" stroke="#9ca3af" tick={{ fontSize: 12 }} />
                <YAxis stroke="#9ca3af" tick={{ fontSize: 12 }} />
                <Tooltip
                  contentStyle={{
                    borderRadius: '8px',
                    border: 'none',
                    boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.1)',
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

          {/* Bar Chart – fica à esquerda na linha de baixo */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200">
            <h3 className="text-lg font-semibold text-gray-800 mb-6">
              Volume por Fila
            </h3>
            <div className="h-80">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={stats.barChartData}
                  layout="vertical"
                  margin={{ top: 10, right: 30, left: 200, bottom: 10 }}
                >
                  <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" />
                  <XAxis type="number" stroke="#9ca3af" tick={{ fontSize: 12 }} />
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
                    <LabelList dataKey="value" position="right" />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Card Próximo Feriado – fica à direita na linha de baixo */}
          {proximoFeriado && (
            <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-200 flex flex-col justify-between">
              <div>
                <h3 className="text-lg font-semibold text-gray-800 mb-4">
                  🗓️ Próximo Feriado
                </h3>

                <p className="text-xl font-bold text-gray-900">
                  {proximoFeriado.nome}
                </p>

                <p className="text-gray-600 text-md mt-1">
                  📅 {proximoFeriado.dataFormatada}
                </p>

                <p className="text-yellow-700 bg-yellow-100 px-4 py-2 mt-4 rounded-lg text-sm border border-yellow-200">
                  ⚠️ Verifique se a operação vai funcionar neste dia.
                </p>
              </div>

              {/* Botão de escala para o próximo feriado – link fixo para todos */}
              <a
                href={LINK_ESCALA_FERIADO}
                target="_blank"
                rel="noopener noreferrer"
                className="inline-block mt-4 bg-purple-600 text-white px-4 py-2 rounded-lg hover:bg-purple-700 transition-colors text-sm self-start"
              >
                🔍 Consultar Escala
              </a>
            </div>
          )}
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
                          {row.data_hora || 'Data não disponível'}
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
                        ? 'Carregando dados do Excel...'
                        : allData.length === 0
                        ? 'Nenhum dado carregado ainda. Verifique o link do Excel publicado ou peça para um supervisor conferir.'
                        : 'Arquivo carregado, mas não há chamadas para este agente e período selecionado.'}
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
