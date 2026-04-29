/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import Papa from 'papaparse';
import { 
  Search, 
  Calendar, 
  RefreshCw, 
  ArrowUpCircle, 
  ArrowDownCircle, 
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Package, 
  AlertCircle,
  ChevronDown,
  ChevronUp,
  ChevronRight,
  ChevronLeft,
  History,
  Filter,
  Download,
  XCircle,
  Eye,
  EyeOff,
  LayoutDashboard,
  AlertTriangle,
  Menu,
  X,
  Home,
  MessageSquare,
  ShieldCheck,
  Zap,
  Ticket,
  Users,
  Settings,
  Activity
} from 'lucide-react';
import { format, parse, isValid, startOfDay, endOfDay, startOfMonth, endOfMonth, eachDayOfInterval, isBefore, isEqual } from 'date-fns';
import * as XLSX from 'xlsx-js-style';
import { cn } from './lib/utils';
import { motion, AnimatePresence } from 'motion/react';

// Types
interface OpeningStockRow {
  partName: string;
  monthlyStocks: Record<string, number>;
}

interface Transaction {
  date: string;
  parsedDate: Date | null;
  partName: string;
  trackingNumber: string;
  type: string;
  quantity: number;
  shift: string;
  department: 'oil-seal' | 'quality' | 'trimming' | 'molding' | 'fg-store' | 'mini-store' | 'bonding' | 'phosphate' | 'auto-clave' | 'extrusion';
}

interface StockSummary {
  itemId: string;
  jobId: string;
  openingStock: number;
  moldIn: number;
  reworkIn: number;
  totalIn: number;
  qcOut: number;
  rejectionOut: number;
  fgReworkIn: number;
  metalStoreIn: number;
  customerRejectionIn: number;
  oilSealTrimmingIn: number;
  trimmingIn: number;
  extrusionIn: number;
  rejectionOutToRps: number;
  metalStoreOut: number;
  oilSealTrimmingOut: number;
  trimmingOut: number;
  fgOut: number;
  extrusionOut: number;
  extrusionProdIn: number;
  extrusionMetalIn: number;
  extrusionMiniStoreIn: number;
  extrusionTrimOut: number;
  trimmingVendorIn: number;
  trimmingQcReworkIn: number;
  trimmingMoldingIn: number;
  trimmingMetalStoreIn: number;
  trimmingExtrusionIn: number;
  trimmingQcOut: number;
  trimmingVendorOut: number;
  trimmingRejectionOutToRps: number;
  // Auto Clave specific
  autoClaveProdIn: number;
  autoClaveMiniStoreIn: number;
  autoClaveMetalIn: number;
  autoClaveReworkIn: number;
  autoClaveRejectionOut: number;
  autoClaveMetalOut: number;
  // Phosphate specific
  phosphateOutToBonding: number;
  // FG Store specific
  qcIn: number;
  customerReturnIn: number;
  autoClaveIn: number;
  qcReworkOut: number;
  // Mini Store specific
  compoundIn: number;
  moldReturnIn: number;
  vendorOut: number;
  injectOut: number;
  oilSealOut: number;
  moldOut: number;
  autoClaveOut: number;
  labOut: number;
  totalOut: number;
  // Bonding specific
  chemicalStoreIn: number;
  phosphateIn: number;
  injcMoldOut: number;
  hvcmOut: number;
  rejectionOutToMetalStore: number;
  vendorOpeningStock: number;
  currentStock: number;
  vendorStock: number;
  totalStock: number;
  nextMonthOpeningStock: number | null;
  hasActivity: boolean;
}

type SortField = keyof StockSummary;
interface SortConfig {
  field: SortField;
  direction: 'asc' | 'desc';
}

const SHEET_ID = '1sBWtgNvzJzXUeYM0uGuSgDACKf7vQnfTuimGDQPrUtM';
const MOLDING_SHEET_ID = '1i-yUEG2VYyMmXUgiFI5qVwMrGLjQfVx0DW6g18yK_yM'; // Updated with provided ID
const QUALITY_SHEET_ID = '16h-bl-eQWb-C4AcK8sMr6yOCfax1psZBDAOTfOisYdA';
const MINI_STORE_SHEET_ID = '10J-GQoPN4wgEUxj9qrCuozZzXIbNq1cB5klKXle5ZR4';
const FG_STORE_SHEET_ID = '1zLy_Te9bReLVV7wSGcoQQw55POlqty9hIUz07jExWeU';
const TRIMMING_SHEET_ID = '1qZOHb3VvOrdxKi0ZWalPZlT_-MGxkrZH8J0enfSz0v0';
const BONDING_SHEET_ID = '1DJ7y8B-BgEluN7TKbxJsWYqYZLcLCEfZaM5kYsxHMp8';
const PHOSPHATE_SHEET_ID = '118OZYrCoweWvGkfSANu7zKG_ClLPG9EPwN5x7GFpD2w';
const AUTO_CLAVE_SHEET_ID = '15kYMNiwwFb_gqUVMqIWZiV6HpFUpYY6hveNFq7PDKZE';
const EXTRUSION_SHEET_ID = '10nIqQgyP8jYcJeEmz6Wkt7j0psKbPYFBKbiVRynpBUI';

const getUrls = (tab: 'molding' | 'oil-seal' | 'quality' | 'trimming' | 'fg-store' | 'mini-store' | 'bonding' | 'phosphate' | 'auto-clave' | 'extrusion') => {
  let id = SHEET_ID;
  if (tab === 'molding') id = MOLDING_SHEET_ID;
  else if (tab === 'quality') id = QUALITY_SHEET_ID;
  else if (tab === 'mini-store') id = MINI_STORE_SHEET_ID;
  else if (tab === 'trimming') id = TRIMMING_SHEET_ID;
  else if (tab === 'fg-store') id = FG_STORE_SHEET_ID;
  else if (tab === 'bonding') id = BONDING_SHEET_ID;
  else if (tab === 'phosphate') id = PHOSPHATE_SHEET_ID || BONDING_SHEET_ID; // Duplicate Bonding for now
  else if (tab === 'auto-clave') id = AUTO_CLAVE_SHEET_ID || BONDING_SHEET_ID; // Duplicate Bonding for now
  else if (tab === 'extrusion') id = EXTRUSION_SHEET_ID;
  
  return {
    OPEN_PCS_URL: tab === 'molding' ? '' : `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=Open%20Pcs`,
    VENDOR_OPEN_PCS_URL: tab === 'trimming' ? `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=Vendor%20Open%20Pcs` : '',
    TRANSACTION_LOG_URL: `https://docs.google.com/spreadsheets/d/${id}/gviz/tq?tqx=out:csv&sheet=Transaction%20Log`
  };
};

const normalizeMonth = (monthStr: string) => {
  if (!monthStr) return '';
  const cleanStr = monthStr.trim();
  
  // Try common formats
  const formats = ['MMM-yy', 'MMM yyyy', 'MMMM yyyy', 'MMM-yyyy', 'MMMM-yy'];
  for (const f of formats) {
    try {
      const parsed = parse(cleanStr, f, new Date());
      if (isValid(parsed)) return format(parsed, 'MMM-yy');
    } catch { continue; }
  }
  
  // Fallback: if it looks like "Apr 2026" or similar, try to extract
  try {
    const date = new Date(cleanStr);
    if (isValid(date)) return format(date, 'MMM-yy');
  } catch {}
  
  return cleanStr;
};

const MENU_ITEMS = [
  { id: 'job-tracking', name: 'JOB TRACKING', icon: <History className="w-4 h-4" /> },
  { id: 'mini-store', name: 'MINI STORE', icon: <Package className="w-4 h-4" /> },
  { id: 'molding', name: 'MOLDING', icon: <Zap className="w-4 h-4" /> },
  { id: 'oil-seal', name: 'OIL SEAL TRIMMING', icon: <ShieldCheck className="w-4 h-4" /> },
  { id: 'trimming', name: 'TRIMMING', icon: <Package className="w-4 h-4" /> },
  { id: 'quality', name: 'QUALITY', icon: <ShieldCheck className="w-4 h-4" /> },
  { id: 'fg-store', name: 'FG STORE', icon: <Package className="w-4 h-4" /> },
  { id: 'bonding', name: 'BONDING', icon: <ShieldCheck className="w-4 h-4" /> },
  { id: 'phosphate', name: 'PHOSPHATE', icon: <ShieldCheck className="w-4 h-4" /> },
  { id: 'auto-clave', name: 'AUTO CLAVE', icon: <ShieldCheck className="w-4 h-4" /> },
  { id: 'extrusion', name: 'EXTRUSION', icon: <Zap className="w-4 h-4" /> },
];

export default function App() {
  const [activeTab, setActiveTab] = useState<'molding' | 'oil-seal' | 'quality' | 'trimming' | 'job-tracking' | 'fg-store' | 'mini-store' | 'bonding' | 'phosphate' | 'auto-clave' | 'extrusion'>('job-tracking');
  const [openingStocks, setOpeningStocks] = useState<OpeningStockRow[]>([]);
  const [vendorOpeningStocks, setVendorOpeningStocks] = useState<OpeningStockRow[]>([]);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [searchTerm, setSearchTerm] = useState('');
  const [filterSearch, setFilterSearch] = useState('');
  const [jobSearchTerm, setJobSearchTerm] = useState('');
  const [showDuplicatesOnly, setShowDuplicatesOnly] = useState(false);
  const [selectedMonth, setSelectedMonth] = useState<string>(''); 
  const [availableMonths, setAvailableMonths] = useState<string[]>([]);
  const [jobColWidth, setJobColWidth] = useState(200);
  const [partColWidth, setPartColWidth] = useState(300);
  const [mainPartColWidth, setMainPartColWidth] = useState(300);
  const [mainJobColWidth, setMainJobColWidth] = useState(150);
  const resizingRef = useRef<{ col: 'job' | 'part' | 'mainPart' | 'mainJob', startX: number, startWidth: number } | null>(null);

  const handleResizeMouseDown = (e: React.MouseEvent, col: 'job' | 'part' | 'mainPart' | 'mainJob') => {
    e.preventDefault();
    resizingRef.current = {
      col,
      startX: e.pageX,
      startWidth: col === 'job' ? jobColWidth : col === 'part' ? partColWidth : col === 'mainPart' ? mainPartColWidth : mainJobColWidth
    };
    document.addEventListener('mousemove', handleResizeMouseMove);
    document.addEventListener('mouseup', handleResizeMouseUp);
    document.body.style.cursor = 'col-resize';
  };

  const handleResizeMouseMove = (e: MouseEvent) => {
    if (!resizingRef.current) return;
    const diff = e.pageX - resizingRef.current.startX;
    const newWidth = Math.max(100, resizingRef.current.startWidth + diff);
    if (resizingRef.current.col === 'job') {
      setJobColWidth(newWidth);
    } else if (resizingRef.current.col === 'part') {
      setPartColWidth(newWidth);
    } else if (resizingRef.current.col === 'mainPart') {
      setMainPartColWidth(newWidth);
    } else if (resizingRef.current.col === 'mainJob') {
      setMainJobColWidth(newWidth);
    }
  };

  const handleResizeMouseUp = () => {
    resizingRef.current = null;
    document.removeEventListener('mousemove', handleResizeMouseMove);
    document.removeEventListener('mouseup', handleResizeMouseUp);
    document.body.style.cursor = '';
  };
  
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [selectedPartNames, setSelectedPartNames] = useState<string[]>([]);
  const [isFilterOpen, setIsFilterOpen] = useState(false);
  const [expandedJob, setExpandedJob] = useState<string | null>(null);
  const [selectedDates, setSelectedDates] = useState<string[]>([]);
  const [isDateFilterOpen, setIsDateFilterOpen] = useState(false);
  const [dateFilterSearch, setDateFilterSearch] = useState('');
  const [sortConfig, setSortConfig] = useState<SortConfig>({ field: 'itemId', direction: 'asc' });
  const [lastUpdated, setLastUpdated] = useState<Date | null>(null);
  const [hideZeroColumns, setHideZeroColumns] = useState(false);
  const [showDailySummary, setShowDailySummary] = useState(false);
  const [showJobSummary, setShowJobSummary] = useState(false);
  const [dailySortConfig, setDailySortConfig] = useState<{ key: string, direction: 'asc' | 'desc' } | null>({ key: 'date', direction: 'desc' });
  const dateCacheMap = useRef<Map<string, Date | null>>(new Map());

  const [showAllJobs, setShowAllJobs] = useState(false);
  const [showJobColumn, setShowJobColumn] = useState(true);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const isQualityTab = activeTab === 'quality' || activeTab === 'mini-store';
  const [dataCache, setDataCache] = useState<Record<string, {
    openingStocks: OpeningStockRow[];
    vendorOpeningStocks?: OpeningStockRow[];
    transactions: Transaction[];
    availableMonths: string[];
  }>>({});

  const resetFilters = () => {
    setSearchTerm('');
    setFilterSearch('');
    setJobSearchTerm('');
    setStartDate('');
    setEndDate('');
    setSelectedPartNames([]);
    setSelectedDates([]);
    setDateFilterSearch('');
    
    // Reset to current month
    if (availableMonths.length > 0) {
      const now = new Date();
      const currentMonthStr = format(now, 'MMM-yy');
      const currentMonthMatch = availableMonths.find(m => m.toLowerCase() === currentMonthStr.toLowerCase());
      
      if (currentMonthMatch) {
        setSelectedMonth(currentMonthMatch);
      } else {
        setSelectedMonth(availableMonths[availableMonths.length - 1]);
      }
    }
  };

  const isColVisible = (total: number) => {
    if (!hideZeroColumns) return true;
    return total !== 0;
  };

  // Fetch Data
  const fetchData = async (tabToFetch: 'molding' | 'oil-seal' | 'quality' | 'trimming' | 'fg-store' | 'mini-store' | 'bonding' | 'phosphate' | 'extrusion', retry: number = 0, force: boolean = false) => {
    const retryCount = retry;
    const isForced = force === true;
    
    // If we have cached data for the tab, use it and return unless forced
    if (dataCache[tabToFetch] && retryCount === 0 && !isForced) {
      if (tabToFetch === activeTab) {
        setOpeningStocks(dataCache[tabToFetch].openingStocks);
        setTransactions(dataCache[tabToFetch].transactions);
        setAvailableMonths(dataCache[tabToFetch].availableMonths);
      }
      return; 
    } 
    
    if (tabToFetch === activeTab) {
      setLoading(true);
    }

    setError(null);
    try {
      const { OPEN_PCS_URL, VENDOR_OPEN_PCS_URL, TRANSACTION_LOG_URL } = getUrls(tabToFetch);
      
      // Fetch Opening Stocks and Transactions in parallel
      const fetchPromises = [];
      if (OPEN_PCS_URL) {
        fetchPromises.push(fetch(OPEN_PCS_URL, { cache: 'no-store' }));
      } else {
        fetchPromises.push(Promise.resolve({ ok: true, text: () => Promise.resolve('') }));
      }

      if (VENDOR_OPEN_PCS_URL) {
        fetchPromises.push(fetch(VENDOR_OPEN_PCS_URL, { cache: 'no-store' }));
      } else {
        fetchPromises.push(Promise.resolve({ ok: true, text: () => Promise.resolve('') }));
      }

      fetchPromises.push(fetch(TRANSACTION_LOG_URL, { cache: 'no-store' }));

      const [openPcsRes, vendorOpenPcsRes, transRes] = await Promise.all(fetchPromises as Promise<any>[]);

      if (!openPcsRes.ok || !transRes.ok || (vendorOpenPcsRes && !vendorOpenPcsRes.ok)) {
        const failedUrl = !openPcsRes.ok ? OPEN_PCS_URL : !transRes.ok ? TRANSACTION_LOG_URL : VENDOR_OPEN_PCS_URL;
        const status = !openPcsRes.ok ? openPcsRes.status : !transRes.ok ? transRes.status : vendorOpenPcsRes.status;
        
        if (status === 404) {
          throw new Error(`Sheet tabs not found in ${tabToFetch} sheet. Please ensure the Google Sheet has tabs named exactly "Open Pcs", "Vendor Open Pcs" (for Trimming), and "Transaction Log".`);
        }
        throw new Error(`Google Sheets returned error ${status} for ${failedUrl}. Please check if the sheet is shared correctly.`);
      }

      const [openPcsText, vendorOpenPcsText, transText] = await Promise.all([
        openPcsRes.text(),
        vendorOpenPcsRes ? vendorOpenPcsRes.text() : Promise.resolve(''),
        transRes.text()
      ]);

      // Check if we got HTML instead of CSV (usually means a redirect to login page)
      if ((openPcsText && openPcsText.includes('<!DOCTYPE html>')) || (vendorOpenPcsText && vendorOpenPcsText.includes('<!DOCTYPE html>')) || transText.includes('<!DOCTYPE html>')) {
        throw new Error('Access denied. Please ensure the Google Sheet is shared as "Anyone with the link can view".');
      }

      const openPcsData = openPcsText ? Papa.parse(openPcsText, { header: false, skipEmptyLines: true }).data as string[][] : [];
      const vendorOpenPcsData = vendorOpenPcsText ? Papa.parse(vendorOpenPcsText, { header: false, skipEmptyLines: true }).data as string[][] : [];
      const transData = Papa.parse(transText, { header: false, skipEmptyLines: true }).data as string[][];

      if (transData.length === 0) {
        throw new Error(`Transaction Log in ${tabToFetch} sheet is empty.`);
      }

      // Dynamically detect column indices from the first row (headers)
      const headerRow = transData[0].map(h => (h || '').trim().toLowerCase());
      
      const findIndexIgnoringCase = (keywords: string[]) => {
        return headerRow.findIndex(h => keywords.some(k => h.includes(k.toLowerCase())));
      };

      const dateIdx = findIndexIgnoringCase(['date', 'dt']);
      const partIdx = findIndexIgnoringCase(['part name', 'description', 'part no', 'item', 'particular']);
      const typeIdx = findIndexIgnoringCase(['transaction type', 'type', 'category', 'status', 'particular']);
      const qtyIdx = findIndexIgnoringCase(['quantity', 'qty', 'pcs', 'count', 'amount']);
      
      // Tab-specific Job # indices as fallback, but try dynamic first
      let jobIdx = findIndexIgnoringCase(['job #', 'job no', 'tracking', 't no', 'serial no', 'card no', 'batch']);
      if (jobIdx === -1) {
        if (tabToFetch === 'trimming') jobIdx = 8; // Column I 
        else if (tabToFetch === 'quality' || tabToFetch === 'mini-store') jobIdx = 6; // Column G
        else if (tabToFetch === 'oil-seal' || tabToFetch === 'bonding' || tabToFetch === 'extrusion') jobIdx = 7; // Column H
        else if (tabToFetch === 'molding') jobIdx = 7; // Column H
      }

      // Valid indices fallback if header detection failed for mandatory columns
      // If we found 'Particular' for both part and type, we need to be careful.
      // Usually, Part is column 1 and Type is column 3/4.
      const finalDateIdx = dateIdx === -1 ? 0 : dateIdx;
      const finalPartIdx = partIdx === -1 ? 1 : partIdx;
      const finalTypeIdx = typeIdx === -1 ? (tabToFetch === 'quality' || tabToFetch === 'mini-store' ? 2 : 3) : typeIdx;
      const finalQtyIdx = qtyIdx === -1 ? (tabToFetch === 'quality' || tabToFetch === 'mini-store' ? 4 : 4) : qtyIdx;
      const finalJobIdx = jobIdx === -1 ? 5 : jobIdx;

      let months: string[] = [];
      let stocks: OpeningStockRow[] = [];
      let vendorStocks: OpeningStockRow[] = [];
      let parsedTransactions: Transaction[] = [];

      // Skip header row (index 0)
      parsedTransactions = transData.slice(1).map((row, rIdx) => {
        const dateStr = (row[finalDateIdx] || '').trim();
        const typeStr = (row[finalTypeIdx] || '').trim();
        const qtyRaw = (row[finalQtyIdx] || '').toString().replace(/,/g, '');
        const qtyVal = parseFloat(qtyRaw);

        return {
          date: dateStr,
          parsedDate: parseSheetDate(dateStr),
          partName: (row[finalPartIdx] || '').trim(),
          trackingNumber: (row[finalJobIdx] || '').trim(),
          type: typeStr,
          quantity: isNaN(qtyVal) ? 0 : qtyVal,
          shift: '', // Fallback
          department: tabToFetch
        };
      }).filter(t => t.partName && t.partName.length > 0 && t.date);

      // Extract months from transactions for ALL tabs
      const transactionMonths = new Set<string>();
      parsedTransactions.forEach(t => {
        if (t.parsedDate) {
          transactionMonths.add(format(t.parsedDate, 'MMM-yy'));
        }
      });

      if (tabToFetch === 'molding') {
        // For molding, we extract part names and months from transactions
        const uniqueParts = Array.from(new Set(parsedTransactions.map(t => t.partName)));
        stocks = uniqueParts.map(part => ({
          partName: part,
          monthlyStocks: {} // No opening stock for molding
        }));

        months = Array.from(transactionMonths).sort((a, b) => {
          const dateA = parse(a, 'MMM-yy', new Date());
          const dateB = parse(b, 'MMM-yy', new Date());
          return dateA.getTime() - dateB.getTime();
        });
      } else if (openPcsData.length > 0) {
        const rawHeaders = openPcsData[0].slice(2).filter(h => h && h.trim() !== '');
        const normalizedHeaders = rawHeaders.map(h => normalizeMonth(h));
        
        // Merge with transaction months to ensure all active months are available
        const combinedMonthsSet = new Set([...normalizedHeaders, ...transactionMonths]);
        months = Array.from(combinedMonthsSet).sort((a, b) => {
          const dateA = parse(a, 'MMM-yy', new Date());
          const dateB = parse(b, 'MMM-yy', new Date());
          return dateA.getTime() - dateB.getTime();
        });

        stocks = openPcsData.slice(1).map(row => {
          const monthlyStocks: Record<string, number> = {};
          rawHeaders.forEach((rawMonth, idx) => {
            const normalized = normalizeMonth(rawMonth);
            const val = parseFloat(row[idx + 2]?.replace(/,/g, '') || '0');
            monthlyStocks[normalized] = isNaN(val) ? 0 : val;
          });
          return {
            partName: (row[0] || '').trim(),
            monthlyStocks
          };
        }).filter(s => s.partName && s.partName.length > 0);

        // Parse Vendor Opening Stocks for Trimming
        if (tabToFetch === 'trimming' && vendorOpenPcsData.length > 0) {
          const rawVHeaders = vendorOpenPcsData[0].slice(2).filter(h => h && h.trim() !== '');
          vendorStocks = vendorOpenPcsData.slice(1).map(row => {
            const monthlyStocks: Record<string, number> = {};
            rawVHeaders.forEach((rawMonth, idx) => {
              const normalized = normalizeMonth(rawMonth);
              const val = parseFloat(row[idx + 2]?.replace(/,/g, '') || '0');
              monthlyStocks[normalized] = isNaN(val) ? 0 : val;
            });
            return {
              partName: (row[0] || '').trim(),
              monthlyStocks
            };
          }).filter(s => s.partName && s.partName.length > 0);
        }
      }

      if (tabToFetch === activeTab) {
        setAvailableMonths(months);
        setOpeningStocks(stocks);
        setVendorOpeningStocks(vendorStocks);
        setTransactions(parsedTransactions);
        setLastUpdated(new Date());

        // Default month selection logic
        if (!selectedMonth && months.length > 0) {
          const now = new Date();
          const currentMonthStr = format(now, 'MMM-yy');
          const match = months.find(m => m.toLowerCase() === currentMonthStr.toLowerCase());
          setSelectedMonth(match || months[months.length - 1]);
        }
      }

      // Update cache
      setDataCache(prev => ({
        ...prev,
        [tabToFetch]: {
          openingStocks: stocks,
          vendorOpeningStocks: vendorStocks,
          transactions: parsedTransactions,
          availableMonths: months
        }
      }));

    } catch (err) {
      console.error(`Error fetching data for ${tabToFetch}:`, err);
      
      if (retryCount < 1) {
        setTimeout(() => fetchData(tabToFetch, retryCount + 1, force), 1000);
        return;
      }

      const message = err instanceof Error ? err.message : 'An unknown error occurred';
      if (tabToFetch === activeTab) {
        setError(message);
      }
    } finally {
      if (tabToFetch === activeTab) {
        setLoading(false);
      }
    }
  };

  // Initial fetch for all tabs
  useEffect(() => {
    const fetchAll = async () => {
      setLoading(true);
      await Promise.all([
        fetchData('molding', 0, true),
        fetchData('oil-seal', 0, true),
        fetchData('bonding', 0, true),
        fetchData('extrusion', 0, true),
        fetchData('quality', 0, true),
        fetchData('trimming', 0, true),
        fetchData('fg-store', 0, true),
        fetchData('mini-store', 0, true)
      ]);
      setLoading(false);
    };
    fetchAll();
  }, []);

  // Sync state when activeTab changes (from cache)
  useEffect(() => {
    if (activeTab === 'job-tracking') {
      // For job tracking, we don't need to sync specific openingStocks/transactions
      // as it uses allTransactions from dataCache
      return;
    }
    if (dataCache[activeTab]) {
      const newMonths = dataCache[activeTab].availableMonths;
      setOpeningStocks(dataCache[activeTab].openingStocks);
      setVendorOpeningStocks(dataCache[activeTab].vendorOpeningStocks || []);
      setTransactions(dataCache[activeTab].transactions);
      setAvailableMonths(newMonths);

      // Ensure selectedMonth is valid for the new tab
      if (selectedMonth && newMonths.length > 0 && !newMonths.includes(selectedMonth)) {
        const now = new Date();
        const currentMonthStr = format(now, 'MMM-yy');
        const match = newMonths.find(m => m.toLowerCase() === currentMonthStr.toLowerCase());
        setSelectedMonth(match || newMonths[newMonths.length - 1]);
      } else if (!selectedMonth && newMonths.length > 0) {
        const now = new Date();
        const currentMonthStr = format(now, 'MMM-yy');
        const match = newMonths.find(m => m.toLowerCase() === currentMonthStr.toLowerCase());
        setSelectedMonth(match || newMonths[newMonths.length - 1]);
      }
    } else {
      // If not in cache for some reason, fetch it
      fetchData(activeTab as 'molding' | 'oil-seal' | 'quality' | 'trimming' | 'bonding' | 'phosphate' | 'fg-store' | 'mini-store' | 'extrusion');
    }
  }, [activeTab, dataCache]);

  // Helper to parse date robustly in local time
  const parseSheetDate = (dateStr: string) => {
    if (!dateStr) return null;
    const cleanStr = dateStr.trim();
    if (dateCacheMap.current.has(cleanStr)) return dateCacheMap.current.get(cleanStr)!;
    
    // Try common formats explicitly to avoid browser-specific new Date() behavior
    const formats = [
      'dd/MM/yyyy', 
      'MM/dd/yyyy', 
      'yyyy-MM-dd', 
      'dd-MM-yyyy', 
      'MMM dd, yyyy',
      'd/M/yyyy',
      'M/d/yyyy',
      'd-M-yyyy',
      'M-d-yyyy',
      'dd-MMM-yy',
      'dd-MMM-yyyy',
      'd-MMM-yy',
      'd-MMM-yyyy'
    ];
    
    let result: Date | null = null;
    for (const f of formats) {
      try {
        const parsed = parse(cleanStr, f, new Date());
        if (isValid(parsed)) {
          result = startOfDay(parsed);
          break;
        }
      } catch { continue; }
    }
    
    if (!result) {
      // Fallback to standard JS date parsing only if formats fail, but wrap in startOfDay
      const fallback = new Date(cleanStr);
      if (isValid(fallback)) result = startOfDay(fallback);
    }
    
    dateCacheMap.current.set(cleanStr, result);
    return result;
  };

  // Unique Dates from Transactions
  const availableDates = useMemo(() => {
    // Group by unique date string first
    const dateMap = new Map<string, Date | null>();
    transactions.forEach(t => {
      if (t.date && !dateMap.has(t.date)) {
        dateMap.set(t.date, t.parsedDate);
      }
    });

    return Array.from(dateMap.keys()).sort((a: string, b: string) => {
      const dateA = dateMap.get(a);
      const dateB = dateMap.get(b);
      if (!dateA || !dateB) return 0;
      return dateB.getTime() - dateA.getTime(); // Newest first
    });
  }, [transactions]);

  const filteredDatesInDropdown = useMemo(() => {
    return availableDates.filter(date => 
      date.toLowerCase().includes(dateFilterSearch.toLowerCase())
    );
  }, [availableDates, dateFilterSearch]);

  const nextMonthName = useMemo(() => {
    if (!selectedMonth || availableMonths.length === 0) return null;
    const currentIndex = availableMonths.indexOf(selectedMonth);
    if (currentIndex !== -1 && currentIndex < availableMonths.length - 1) {
      return availableMonths[currentIndex + 1];
    }
    return null;
  }, [selectedMonth, availableMonths]);

  // Calculate all items with their activity for the current period
  const allItemsWithActivity = useMemo(() => {
    if (!selectedMonth && !startDate && selectedDates.length === 0) return [];

    // 1. Pre-calculate date filter parameters to avoid repeated parsing
    const start = startDate ? startOfDay(parse(startDate, 'yyyy-MM-dd', new Date())) : null;
    const end = endDate ? endOfDay(parse(endDate, 'yyyy-MM-dd', new Date())) : null;
    
    let headerDate: Date | null = null;
    if (selectedMonth) {
      headerDate = parse(selectedMonth, 'MMM-yy', new Date());
      if (!isValid(headerDate)) {
        headerDate = parse(selectedMonth, 'MMM yyyy', new Date());
      }
    }

    // 2. Filter transactions by date ONCE
    const dateFilteredTransactions = transactions.filter(t => {
      // Ensure we only process transactions for the active department
      if (t.department && t.department !== activeTab) return false;

      const tDate = t.parsedDate;
      if (!tDate || !isValid(tDate)) return false;

      if (selectedDates.length > 0) {
        return selectedDates.includes(t.date);
      }

      if (start || end) {
        if (start && tDate < start) return false;
        if (end && tDate > end) return false;
        return true;
      }

      if (headerDate && isValid(headerDate)) {
        return tDate.getMonth() === headerDate.getMonth() && tDate.getFullYear() === headerDate.getFullYear();
      }
      
      if (selectedMonth) {
        const monthName = format(tDate, 'MMMM');
        return selectedMonth.toLowerCase().includes(monthName.toLowerCase().substring(0, 3));
      }

      return false;
    });

    // 3. Group transactions by part name or job number for O(1) lookup
    const transGroups = new Map<string, Transaction[]>();
    dateFilteredTransactions.forEach(t => {
      const key = showJobSummary 
        ? `${(t.trackingNumber || 'No Job').trim()}|${t.partName.trim()}`
        : t.partName.toString().trim().toLowerCase();
      
      if (!transGroups.has(key)) transGroups.set(key, []);
      transGroups.get(key)!.push(t);
    });

    // 3.1 Group vendor opening stocks by part name
    const vendorStocksByPart = new Map<string, number>();
    if (activeTab === 'trimming') {
      vendorOpeningStocks.forEach(vs => {
        const key = vs.partName.toString().trim().toLowerCase();
        const val = vs.monthlyStocks[selectedMonth] || 0;
        vendorStocksByPart.set(key, val);
      });
    }

    // 4. Calculate items based on grouping
    if (showJobSummary) {
      // Job-wise summary
      const jobItems = Array.from(transGroups.entries()).map(([key, filteredTrans]) => {
        const [jobId, pName] = key.split('|');
        const sPartKey = pName.toLowerCase();
        
        // For Job Summary, opening stock is 0 for now as we don't have job-level opening data
        const opening = 0;
        const vendorOpening = 0;
        const nextMonthOpening = 0;

        // Oil Seal specific
        let moldIn = 0, reworkIn = 0, qcOut = 0, rejectionOut = 0;
        // Extrusion specific
        let extrusionProdIn = 0, extrusionMetalIn = 0, extrusionMiniStoreIn = 0, extrusionTrimOut = 0;
        // Bonding specific
        let chemicalStoreIn = 0, phosphateIn = 0, injcMoldOut = 0, hvcmOut = 0, rejectionOutToMetalStore = 0;
        // Phosphate specific
        let phosphateOutToBonding = 0;
        // Auto Clave specific
        let autoClaveProdIn = 0, autoClaveMiniStoreIn = 0, autoClaveMetalIn = 0, autoClaveRejectionOut = 0, autoClaveMetalOut = 0, autoClaveReworkIn = 0;
        // Quality specific
        let fgReworkIn = 0, metalStoreIn = 0, customerRejectionIn = 0, oilSealTrimmingIn = 0, trimmingIn = 0, extrusionIn = 0;
        let rejectionOutToRps = 0, metalStoreOut = 0, oilSealTrimmingOut = 0, trimmingOut = 0, fgOut = 0, extrusionOut = 0;
        // Trimming specific
        let trimmingVendorIn = 0, trimmingQcReworkIn = 0, trimmingMoldingIn = 0, trimmingMetalStoreIn = 0, trimmingExtrusionIn = 0;
        let trimmingQcOut = 0, trimmingVendorOut = 0, trimmingRejectionOutToRps = 0;
        // FG Store specific
        let qcIn = 0, autoClaveIn = 0, qcReworkOut = 0;
        // Mini Store specific
        let compoundIn = 0, moldReturnIn = 0, vendorOut = 0, injectOut = 0, oilSealOut = 0, moldOut = 0, autoClaveOut = 0, labOut = 0;
        // Totals
        let totalIn = 0, totalOut = 0;

        filteredTrans.forEach(t => {
          const type = t.type.toLowerCase();
          const qty = t.quantity;

          if (activeTab === 'bonding') {
            if (type.includes('metal') && type.includes('in')) metalStoreIn += qty;
            else if (type.includes('chemical') && type.includes('in')) chemicalStoreIn += qty;
            else if (type.includes('phosphate') && type.includes('in')) phosphateIn += qty;
            else if (type.includes('mold') && type.includes('in')) moldIn += qty;
            else if (type.includes('injc') && type.includes('out')) injcMoldOut += qty;
            else if (type.includes('oil') && type.includes('seal') && type.includes('out')) oilSealOut += qty;
            else if (type.includes('hvcm') && type.includes('out')) hvcmOut += qty;
            else if (type.includes('rejection') && type.includes('out')) rejectionOutToMetalStore += qty;
            else if (type.includes('rejection') && !type.includes('in')) rejectionOutToMetalStore += qty;
          } else if (activeTab === 'auto-clave') {
            if (type.includes('prod') && type.includes('in')) autoClaveProdIn += qty;
            else if (type.includes('mini') && type.includes('store') && type.includes('in')) autoClaveMiniStoreIn += qty;
            else if (type.includes('metal') && type.includes('in')) autoClaveMetalIn += qty;
            else if (type.includes('rework') && type.includes('in')) autoClaveReworkIn += qty;
            else if (type.includes('metal') && type.includes('out')) autoClaveMetalOut += qty;
            else if (type.includes('reject') && type.includes('rps')) autoClaveRejectionOut += qty;
          } else if (activeTab === 'phosphate') {
            if (type.includes('metal') && type.includes('in')) metalStoreIn += qty;
            else if (type.includes('chemical') && type.includes('in')) chemicalStoreIn += qty;
            else if (type.includes('bonding') && type.includes('out')) phosphateOutToBonding += qty;
            else if (type.includes('reject') && type.includes('rps')) rejectionOutToRps += qty;
          } else if (activeTab === 'oil-seal') {
            if (type.includes('mold') && type.includes('in')) moldIn += qty;
            else if (type.includes('re work in') || type.includes('rework in')) reworkIn += qty;
            else if (type.includes('qc out')) qcOut += qty;
            else if (type.includes('rejection to rps out') || type.includes('rejection')) rejectionOut += qty;
          } else if (activeTab === 'extrusion') {
            if ((type.includes('prod') || type.includes('production')) && type.includes('in')) extrusionProdIn += qty;
            else if (type.includes('re work in') || type.includes('rework in')) reworkIn += qty;
            else if (type.includes('metal') && type.includes('in')) extrusionMetalIn += qty;
            else if (type.includes('mini') && type.includes('store') && type.includes('in')) extrusionMiniStoreIn += qty;
            else if (type.includes('reject') && type.includes('rps')) rejectionOutToRps += qty;
            else if (type.includes('fg') && type.includes('out')) fgOut += qty;
            else if (type.includes('trim') && type.includes('out')) extrusionTrimOut += qty;
            else if (type.includes('qc') && type.includes('out')) qcOut += qty;
          } else if (activeTab === 'molding') {
            if (type.includes('rejection out to rps')) rejectionOutToRps += qty;
            else if (type.includes('oil seal trimming out')) oilSealTrimmingOut += qty;
            else if (type.includes('trimming out')) trimmingOut += qty;
          } else if (activeTab === 'quality') {
            const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
            const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

            if (type.includes('fg rework') && isRecd) fgReworkIn += qty;
            else if (type.includes('metal') && isRecd) metalStoreIn += qty;
            else if (type.includes('customer') && isRecd) customerRejectionIn += qty;
            else if (((type.includes('oil') && type.includes('seal')) || type.includes('bonding')) && isRecd) oilSealTrimmingIn += qty;
            else if (type.includes('trim') && isRecd) trimmingIn += qty;
            else if (type.includes('extru') && isRecd) extrusionIn += qty;
            else if (type.includes('reject') && (isSent || !isRecd)) rejectionOutToRps += qty;
            else if (type.includes('metal') && isSent) metalStoreOut += qty;
            else if (((type.includes('oil') && type.includes('seal')) || type.includes('bonding')) && isSent) oilSealTrimmingOut += qty;
            else if (type.includes('trim') && isSent) trimmingOut += qty;
            else if (type.includes('fg') && isSent) fgOut += qty;
            else if (type.includes('extru') && isSent) extrusionOut += qty;

            if (isRecd) totalIn += qty;
            else if (isSent || type.includes('reject')) totalOut += qty;
          } else if (activeTab === 'mini-store') {
            const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
            const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

            if (type.includes('compound') && isRecd) compoundIn += qty;
            else if (type.includes('mold return') && isRecd) moldReturnIn += qty;
            else if (type.includes('vendor') && isSent) vendorOut += qty;
            else if (type.includes('inject') && isSent) injectOut += qty;
            else if ((type.includes('oil seal') || type.includes('bonding')) && isSent) oilSealOut += qty;
            else if (type.includes('mold') && isSent) moldOut += qty;
            else if (type.includes('extru') && isSent) extrusionOut += qty;
            else if (type.includes('auto') && type.includes('clave') && isSent) autoClaveOut += qty;
            else if (type.includes('lab') && isSent) labOut += qty;
            else if (type.includes('reject') && (isSent || !isRecd)) rejectionOutToRps += qty;

            if (isRecd) totalIn += qty;
            else if (isSent || type.includes('reject')) totalOut += qty;
          } else if (activeTab === 'fg-store') {
            const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
            const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

            if (type.includes('customer') && isRecd) customerRejectionIn += qty;
            else if (type.includes('qc') && isRecd) qcIn += qty;
            else if (type.includes('auto') && type.includes('clave') && isRecd) autoClaveIn += qty;
            else if (type.includes('reject') && type.includes('rps') && isSent) rejectionOutToRps += qty;
            else if (type.includes('qc') && type.includes('rework') && isSent) qcReworkOut += qty;
            else if (type.includes('fg') && isSent) fgOut += qty;

            if (isRecd) totalIn += qty;
            else if (isSent) totalOut += qty;
          } else if (activeTab === 'trimming') {
            const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
            const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

            if (type.includes('vendor') && isRecd) trimmingVendorIn += qty;
            else if (type.includes('qc rework')) trimmingQcReworkIn += qty;
            else if (type.includes('mold') && isRecd) trimmingMoldingIn += qty;
            else if (type.includes('metal') && isRecd) trimmingMetalStoreIn += qty;
            else if (type.includes('extru') && isRecd) trimmingExtrusionIn += qty;
            else if (type.includes('qc') && isSent) trimmingQcOut += qty;
            else if (type.includes('vendor') && isSent) trimmingVendorOut += qty;
            else if (type.includes('reject') && (isSent || !isRecd)) trimmingRejectionOutToRps += qty;

            if (isRecd) totalIn += qty;
            else if (isSent || type.includes('reject')) totalOut += qty;
          }
        });

        if (activeTab === 'bonding') {
          totalIn = opening + metalStoreIn + chemicalStoreIn + phosphateIn + moldIn;
          totalOut = injcMoldOut + oilSealOut + hvcmOut + rejectionOutToMetalStore;
        } else if (activeTab === 'auto-clave') {
          totalIn = opening + autoClaveProdIn + autoClaveMiniStoreIn + autoClaveMetalIn + autoClaveReworkIn;
          totalOut = autoClaveRejectionOut + autoClaveMetalOut;
        } else if (activeTab === 'phosphate') {
          totalIn = opening + metalStoreIn + chemicalStoreIn;
          totalOut = phosphateOutToBonding + rejectionOutToRps;
        } else if (activeTab === 'oil-seal') {
          totalIn = opening + moldIn + reworkIn;
          totalOut = qcOut + rejectionOut;
        } else if (activeTab === 'extrusion') {
          totalIn = opening + reworkIn + extrusionProdIn + extrusionMetalIn + extrusionMiniStoreIn;
          totalOut = rejectionOutToRps + fgOut + extrusionTrimOut + qcOut;
        } else if (activeTab === 'molding') {
          totalIn = 0;
          totalOut = rejectionOutToRps + oilSealTrimmingOut + trimmingOut;
        } else if (activeTab === 'trimming') {
          totalIn = opening + trimmingVendorIn + trimmingQcReworkIn + trimmingMoldingIn + trimmingMetalStoreIn + trimmingExtrusionIn;
          totalOut = trimmingQcOut + trimmingVendorOut + trimmingRejectionOutToRps;
        } else if (activeTab === 'quality') {
          totalIn = opening + totalIn;
        } else if (activeTab === 'mini-store') {
          totalIn = opening + compoundIn + moldReturnIn;
          totalOut = vendorOut + injectOut + oilSealOut + moldOut + extrusionOut + autoClaveOut + labOut + rejectionOutToRps;
        } else if (activeTab === 'fg-store') {
          totalIn = opening + totalIn;
        }
        
        const currentStock = totalIn - totalOut;
        const vendorStock = activeTab === 'trimming' ? (vendorOpening + trimmingVendorOut - trimmingVendorIn) : 0;
        const totalStock = activeTab === 'trimming' ? (currentStock + vendorStock) : currentStock;

        return {
          itemId: jobId,
          partName: pName,
          jobId: jobId,
          openingStock: opening,
          vendorOpeningStock: vendorOpening,
          autoClaveProdIn, autoClaveMiniStoreIn, autoClaveMetalIn, autoClaveRejectionOut, autoClaveMetalOut, autoClaveReworkIn,
          fromBondingIn: 0, fromMoldingIn: 0, toQualityOut: 0, toFgStoreOut: 0, toMiniStoreOut: 0,
          moldIn, reworkIn, qcOut, rejectionOut,
          extrusionProdIn, extrusionMetalIn, extrusionMiniStoreIn, extrusionTrimOut,
          chemicalStoreIn, phosphateIn, injcMoldOut, hvcmOut, rejectionOutToMetalStore, phosphateOutToBonding,
          fgReworkIn, metalStoreIn, customerRejectionIn, oilSealTrimmingIn, trimmingIn, extrusionIn,
          rejectionOutToRps, metalStoreOut, oilSealTrimmingOut, trimmingOut, fgOut, extrusionOut,
          trimmingVendorIn, trimmingQcReworkIn, trimmingMoldingIn, trimmingMetalStoreIn, trimmingExtrusionIn,
          trimmingQcOut, trimmingVendorOut, trimmingRejectionOutToRps,
          qcIn, autoClaveIn, qcReworkOut,
          compoundIn, moldReturnIn, vendorOut, injectOut, oilSealOut, moldOut, autoClaveOut, labOut,
          totalIn, totalOut, currentStock,
          vendorStock, totalStock,
          nextMonthOpeningStock: nextMonthOpening,
          hasActivity: totalIn !== 0 || totalOut !== 0
        };
      });
      return jobItems;
    }

    // Default Part-wise summary - COMBINE opening stock list and transaction list to ensure NO DATA IS MISSING
    const allUniquePartKeys = new Set<string>();
    openingStocks.forEach(s => allUniquePartKeys.add(s.partName.toString().trim().toLowerCase()));
    if (!showJobSummary) {
      transGroups.forEach((_, key) => allUniquePartKeys.add(key));
    }

    const partialMatchCache = new Map<string, Transaction[]>();
    const uniqueTransKeys = Array.from(transGroups.keys());

    const items = Array.from(allUniquePartKeys).map(pKey => {
      const stock = openingStocks.find(s => s.partName.toString().trim().toLowerCase() === pKey);
      const opening = activeTab === 'molding' ? 0 : (stock?.monthlyStocks[selectedMonth] || 0);
      const vendorOpening = activeTab === 'trimming' ? (vendorStocksByPart.get(pKey) || 0) : 0;
      const nextMonthOpening = activeTab === 'molding' ? null : (nextMonthName ? (stock?.monthlyStocks[nextMonthName] || 0) : null);
      
      const pName = stock ? stock.partName : (transGroups.get(pKey)?.[0]?.partName || pKey.toUpperCase());

      // Get transactions for this part (O(1) lookup)
      let filteredTrans = transGroups.get(pKey) || [];
      
      // If no exact match, try partial match (optimized)
      if (filteredTrans.length === 0) {
        if (partialMatchCache.has(pKey)) {
          filteredTrans = partialMatchCache.get(pKey)!;
        } else {
          const matchedKeys = uniqueTransKeys.filter(tKey => 
            tKey.includes(pKey) || pKey.includes(tKey)
          );
          filteredTrans = matchedKeys.flatMap(k => transGroups.get(k) || []);
          partialMatchCache.set(pKey, filteredTrans);
        }
      }

      // Oil Seal specific
      let moldIn = 0, reworkIn = 0, qcOut = 0, rejectionOut = 0;
      // Extrusion specific
      let extrusionProdIn = 0, extrusionMetalIn = 0, extrusionMiniStoreIn = 0, extrusionTrimOut = 0;
      // Bonding specific
      let chemicalStoreIn = 0, phosphateIn = 0, injcMoldOut = 0, hvcmOut = 0, rejectionOutToMetalStore = 0;
      // Phosphate specific
      let phosphateOutToBonding = 0;
      // Auto Clave specific
      let autoClaveProdIn = 0, autoClaveMiniStoreIn = 0, autoClaveMetalIn = 0, autoClaveRejectionOut = 0, autoClaveMetalOut = 0, autoClaveReworkIn = 0;
      // Quality specific
      let fgReworkIn = 0, metalStoreIn = 0, customerRejectionIn = 0, oilSealTrimmingIn = 0, trimmingIn = 0, extrusionIn = 0;
      let rejectionOutToRps = 0, metalStoreOut = 0, oilSealTrimmingOut = 0, trimmingOut = 0, fgOut = 0, extrusionOut = 0;
      // Trimming specific
      let trimmingVendorIn = 0, trimmingQcReworkIn = 0, trimmingMoldingIn = 0, trimmingMetalStoreIn = 0, trimmingExtrusionIn = 0;
      let trimmingQcOut = 0, trimmingVendorOut = 0, trimmingRejectionOutToRps = 0;
      // FG Store specific
      let qcIn = 0, autoClaveIn = 0, qcReworkOut = 0;
      // Mini Store specific
      let compoundIn = 0, moldReturnIn = 0, vendorOut = 0, injectOut = 0, oilSealOut = 0, moldOut = 0, autoClaveOut = 0, labOut = 0;
      // Totals
      let totalIn = 0, totalOut = 0;

      filteredTrans.forEach(t => {
        const type = t.type.toLowerCase();
        const qty = t.quantity;

        if (activeTab === 'bonding') {
          if (type.includes('metal') && type.includes('in')) metalStoreIn += qty;
          else if (type.includes('chemical') && type.includes('in')) chemicalStoreIn += qty;
          else if (type.includes('phosphate') && type.includes('in')) phosphateIn += qty;
          else if (type.includes('mold') && type.includes('in')) moldIn += qty;
          else if (type.includes('injc') && type.includes('mold') && type.includes('out')) injcMoldOut += qty;
          else if (type.includes('oil') && type.includes('seal') && type.includes('out')) oilSealOut += qty;
          else if (type.includes('hvcm') && type.includes('out')) hvcmOut += qty;
          else if (type.includes('rejection') && type.includes('out')) rejectionOutToMetalStore += qty;
        } else if (activeTab === 'auto-clave') {
          if (type.includes('prod') && type.includes('in')) autoClaveProdIn += qty;
          else if (type.includes('mini') && type.includes('store') && type.includes('in')) autoClaveMiniStoreIn += qty;
          else if (type.includes('metal') && type.includes('in')) autoClaveMetalIn += qty;
          else if (type.includes('rework') && type.includes('in')) autoClaveReworkIn += qty;
          else if (type.includes('metal') && type.includes('out')) autoClaveMetalOut += qty;
          else if (type.includes('reject') && type.includes('rps')) autoClaveRejectionOut += qty;
        } else if (activeTab === 'phosphate') {
          if (type.includes('metal') && type.includes('in')) metalStoreIn += qty;
          else if (type.includes('chemical') && type.includes('in')) chemicalStoreIn += qty;
          else if (type.includes('bonding') && type.includes('out')) phosphateOutToBonding += qty;
          else if (type.includes('reject') && type.includes('rps')) rejectionOutToRps += qty;
        } else if (activeTab === 'oil-seal') {
          if (type.includes('mold') && type.includes('in')) moldIn += qty;
          else if (type.includes('re work in') || type.includes('rework in')) reworkIn += qty;
          else if (type.includes('qc out')) qcOut += qty;
          else if (type.includes('rejection to rps out') || type.includes('rejection')) rejectionOut += qty;
        } else if (activeTab === 'extrusion') {
          if ((type.includes('prod') || type.includes('production')) && type.includes('in')) { extrusionProdIn += qty; totalIn += qty; }
          else if (type.includes('re work in') || type.includes('rework in')) { reworkIn += qty; totalIn += qty; }
          else if (type.includes('metal') && type.includes('in')) { extrusionMetalIn += qty; totalIn += qty; }
          else if (type.includes('mini') && type.includes('store') && type.includes('in')) { extrusionMiniStoreIn += qty; totalIn += qty; }
          else if (type.includes('reject') && type.includes('rps')) { rejectionOutToRps += qty; totalOut += qty; }
          else if (type.includes('fg') && type.includes('out')) { fgOut += qty; totalOut += qty; }
          else if (type.includes('trim') && type.includes('out')) { extrusionTrimOut += qty; totalOut += qty; }
          else if (type.includes('qc') && type.includes('out')) { qcOut += qty; totalOut += qty; }
        } else if (activeTab === 'molding') {
          // Molding specific types as requested
          if (type.includes('rejection out to rps')) rejectionOutToRps += qty;
          else if (type.includes('oil seal trimming out')) oilSealTrimmingOut += qty;
          else if (type.includes('trimming out')) trimmingOut += qty;
        } else if (activeTab === 'quality') {
          // Quality - improved classification
          const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
          const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

          if (type.includes('fg rework') && isRecd) fgReworkIn += qty;
          else if (type.includes('metal') && isRecd) metalStoreIn += qty;
          else if (type.includes('customer') && isRecd) customerRejectionIn += qty;
          else if (((type.includes('oil') && type.includes('seal')) || type.includes('bonding')) && isRecd) oilSealTrimmingIn += qty;
          else if (type.includes('trim') && isRecd) trimmingIn += qty;
          else if (type.includes('extru') && isRecd) extrusionIn += qty;
          else if (type.includes('reject') && (isSent || !isRecd)) rejectionOutToRps += qty;
          else if (type.includes('metal') && isSent) metalStoreOut += qty;
          else if (((type.includes('oil') && type.includes('seal')) || type.includes('bonding')) && isSent) oilSealTrimmingOut += qty;
          else if (type.includes('trim') && isSent) trimmingOut += qty;
          else if (type.includes('fg') && isSent) fgOut += qty;
          else if (type.includes('extru') && isSent) extrusionOut += qty;

          if (isRecd) totalIn += qty;
          else if (isSent || type.includes('reject')) totalOut += qty;
        } else if (activeTab === 'mini-store') {
          // Mini Store - improved classification
          const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
          const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

          if (type.includes('compound') && isRecd) compoundIn += qty;
          else if (type.includes('mold return') && isRecd) moldReturnIn += qty;
          else if (type.includes('vendor') && isSent) vendorOut += qty;
          else if (type.includes('inject') && isSent) injectOut += qty;
          else if ((type.includes('oil seal') || type.includes('bonding')) && isSent) oilSealOut += qty;
          else if (type.includes('mold') && isSent) moldOut += qty;
          else if (type.includes('extru') && isSent) extrusionOut += qty;
          else if (type.includes('auto') && type.includes('clave') && isSent) autoClaveOut += qty;
          else if (type.includes('lab') && isSent) labOut += qty;
          else if (type.includes('reject') && (isSent || !isRecd)) rejectionOutToRps += qty;

          if (isRecd) totalIn += qty;
          else if (isSent || type.includes('reject')) totalOut += qty;
        } else if (activeTab === 'fg-store') {
          const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
          const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

          if (type.includes('customer') && isRecd) customerRejectionIn += qty;
          else if (type.includes('qc') && isRecd) qcIn += qty;
          else if (type.includes('auto') && type.includes('clave') && isRecd) autoClaveIn += qty;
          else if (type.includes('reject') && type.includes('rps') && isSent) rejectionOutToRps += qty;
          else if (type.includes('qc') && type.includes('rework') && isSent) qcReworkOut += qty;
          else if (type.includes('fg') && isSent) fgOut += qty;

          if (isRecd) totalIn += qty;
          else if (isSent) totalOut += qty;
        } else if (activeTab === 'trimming') {
          // Trimming - improved classification
          const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
          const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

          if (type.includes('vendor') && isRecd) trimmingVendorIn += qty;
          else if (type.includes('qc rework')) trimmingQcReworkIn += qty;
          else if (type.includes('mold') && isRecd) trimmingMoldingIn += qty;
          else if (type.includes('metal store') && isRecd) trimmingMetalStoreIn += qty;
          else if (type.includes('extru') && isRecd) trimmingExtrusionIn += qty;
          else if (type.includes('qc') && isSent) trimmingQcOut += qty;
          else if (type.includes('vendor') && isSent) trimmingVendorOut += qty;
          else if (type.includes('reject') && (isSent || !isRecd)) trimmingRejectionOutToRps += qty;

          if (isRecd) totalIn += qty;
          else if (isSent || type.includes('reject')) totalOut += qty;
        }
      });

      if (activeTab === 'bonding') {
        totalIn = opening + metalStoreIn + chemicalStoreIn + phosphateIn + moldIn;
        totalOut = injcMoldOut + oilSealOut + hvcmOut + rejectionOutToMetalStore;
      } else if (activeTab === 'auto-clave') {
        totalIn = opening + autoClaveProdIn + autoClaveMiniStoreIn + autoClaveMetalIn + autoClaveReworkIn;
        totalOut = autoClaveRejectionOut + autoClaveMetalOut;
      } else if (activeTab === 'phosphate') {
        totalIn = opening + metalStoreIn + chemicalStoreIn;
        totalOut = phosphateOutToBonding + rejectionOutToRps;
      } else if (activeTab === 'oil-seal') {
        totalIn = opening + moldIn + reworkIn;
        totalOut = qcOut + rejectionOut;
      } else if (activeTab === 'extrusion') {
        totalIn = opening + reworkIn + extrusionProdIn + extrusionMetalIn + extrusionMiniStoreIn;
        totalOut = rejectionOutToRps + fgOut + extrusionTrimOut + qcOut;
      } else if (activeTab === 'molding') {
        totalIn = 0;
        totalOut = rejectionOutToRps + oilSealTrimmingOut + trimmingOut;
      } else if (activeTab === 'trimming') {
        totalIn = opening + trimmingVendorIn + trimmingQcReworkIn + trimmingMoldingIn + trimmingMetalStoreIn + trimmingExtrusionIn;
        totalOut = trimmingQcOut + trimmingVendorOut + trimmingRejectionOutToRps;
      } else if (activeTab === 'quality') {
        totalIn = opening + totalIn;
      } else if (activeTab === 'mini-store') {
        totalIn = opening + compoundIn + moldReturnIn;
        totalOut = vendorOut + injectOut + oilSealOut + moldOut + extrusionOut + autoClaveOut + labOut + rejectionOutToRps;
      } else if (activeTab === 'fg-store') {
        totalIn = opening + totalIn;
      }
      
      const currentStock = totalIn - totalOut;
      const vendorStock = activeTab === 'trimming' ? (vendorOpening + trimmingVendorOut - trimmingVendorIn) : 0;
      const totalStock = activeTab === 'trimming' ? (currentStock + vendorStock) : currentStock;

      // Extract Job # from transactions
      const jobIds = Array.from(new Set(filteredTrans.map(t => t.trackingNumber).filter(j => j && j !== '')));
      const jobId = jobIds.join(', ');

      return {
        itemId: pKey,
        partName: pName,
        jobId,
        openingStock: opening || 0,
        vendorOpeningStock: vendorOpening,
        autoClaveProdIn, autoClaveMiniStoreIn, autoClaveMetalIn, autoClaveRejectionOut, autoClaveMetalOut, autoClaveReworkIn,
        fromBondingIn: 0, fromMoldingIn: 0, toQualityOut: 0, toFgStoreOut: 0, toMiniStoreOut: 0,
        moldIn, reworkIn, qcOut, rejectionOut,
        extrusionProdIn, extrusionMetalIn, extrusionMiniStoreIn, extrusionTrimOut,
        chemicalStoreIn, phosphateIn, injcMoldOut, hvcmOut, rejectionOutToMetalStore, phosphateOutToBonding,
        fgReworkIn, metalStoreIn, customerRejectionIn, oilSealTrimmingIn, trimmingIn, extrusionIn,
        rejectionOutToRps, metalStoreOut, oilSealTrimmingOut, trimmingOut, fgOut, extrusionOut,
        trimmingVendorIn, trimmingQcReworkIn, trimmingMoldingIn, trimmingMetalStoreIn, trimmingExtrusionIn,
        trimmingQcOut, trimmingVendorOut, trimmingRejectionOutToRps,
        qcIn, autoClaveIn, qcReworkOut,
        compoundIn, moldReturnIn, vendorOut, injectOut, oilSealOut, moldOut, autoClaveOut, labOut,
        totalIn, totalOut, currentStock,
        vendorStock, totalStock,
        nextMonthOpeningStock: nextMonthOpening,
        hasActivity: (opening || 0) !== 0 || totalIn !== 0 || totalOut !== 0 || (nextMonthOpening || 0) !== 0
      };
    });
    return items;
  }, [openingStocks, transactions, selectedMonth, startDate, endDate, selectedDates, activeTab, availableMonths, nextMonthName, showJobSummary, vendorOpeningStocks]);
  
  const hasAnyNextMonthStock = useMemo(() => {
    if (!nextMonthName) return false;
    return allItemsWithActivity.some(item => (item.nextMonthOpeningStock || 0) > 0);
  }, [allItemsWithActivity, nextMonthName]);

  // Calculate Summary (Final filtered list for the table)
  const dailySummary = useMemo(() => {
    if (!selectedMonth || openingStocks.length === 0 || transactions.length === 0) return [];

    let parsedMonth: Date;
    try {
      parsedMonth = parse(selectedMonth, 'MMM-yy', new Date());
      if (!isValid(parsedMonth)) {
        parsedMonth = parse(selectedMonth, 'MMM yyyy', new Date());
      }
    } catch (e) {
      return [];
    }

    if (!isValid(parsedMonth)) return [];

    const startOfM = startOfMonth(parsedMonth);
    const endOfM = endOfMonth(parsedMonth);
    const allDaysInMonth = eachDayOfInterval({ start: startOfM, end: endOfM });

    // 1. Identify days with transactions in the selected month
    const daysWithTransactions = new Set<string>();
    const monthTrans = transactions.filter(t => {
      // Ensure we only process transactions for the active department
      if (t.department && t.department !== activeTab) return false;

      if (!t.parsedDate) return false;
      
      // Use parsedMonth to check if transaction belongs to the selected month/year
      const isSameMonth = t.parsedDate.getMonth() === parsedMonth.getMonth() && 
                         t.parsedDate.getFullYear() === parsedMonth.getFullYear();
      
      if (isSameMonth) {
        daysWithTransactions.add(format(t.parsedDate, 'yyyy-MM-dd'));
        return true;
      }
      return false;
    });

    const sortedActiveDays = allDaysInMonth.filter(day => 
      daysWithTransactions.has(format(day, 'yyyy-MM-dd'))
    );

    if (sortedActiveDays.length === 0) return [];

    // 2. Group transactions by part name for the WHOLE month
    const transByPart = new Map<string, Transaction[]>();
    monthTrans.forEach(t => {
      const key = t.partName.toString().trim().toLowerCase();
      if (!transByPart.has(key)) transByPart.set(key, []);
      transByPart.get(key)!.push(t);
    });

    const uniqueTransPartNames = Array.from(transByPart.keys());
    const partialMatchCache = new Map<string, string[]>();

    // 3. Calculate totals for each active day
    return sortedActiveDays.map(day => {
      const startOfD = startOfDay(day);
      const endOfD = endOfDay(day);

      let dayTotals: any = {
        date: day,
        openingStock: 0, currentStock: 0,
        moldIn: 0, reworkIn: 0, qcOut: 0, rejectionOut: 0,
        extrusionProdIn: 0, extrusionMetalIn: 0, extrusionMiniStoreIn: 0, extrusionTrimOut: 0,
        chemicalStoreIn: 0, phosphateIn: 0, injcMoldOut: 0, hvcmOut: 0, rejectionOutToMetalStore: 0,
        fgReworkIn: 0, metalStoreIn: 0, customerRejectionIn: 0, oilSealTrimmingIn: 0, trimmingIn: 0, extrusionIn: 0,
        compoundIn: 0, moldReturnIn: 0, vendorOut: 0, injectOut: 0, oilSealOut: 0, moldOut: 0, autoClaveOut: 0, labOut: 0,
        rejectionOutToRps: 0, metalStoreOut: 0, oilSealTrimmingOut: 0, trimmingOut: 0, fgOut: 0, extrusionOut: 0,
        trimmingVendorIn: 0, trimmingQcReworkIn: 0, trimmingMoldingIn: 0, trimmingMetalStoreIn: 0, trimmingExtrusionIn: 0,
        trimmingQcOut: 0, trimmingVendorOut: 0, trimmingRejectionOutToRps: 0,
        autoClaveProdIn: 0, autoClaveMiniStoreIn: 0, autoClaveMetalIn: 0, autoClaveRejectionOut: 0, autoClaveMetalOut: 0, autoClaveReworkIn: 0,
        phosphateOutToBonding: 0,
        qcIn: 0, autoClaveIn: 0, qcReworkOut: 0,
        totalIn: 0, totalOut: 0
      };

      openingStocks.forEach(stock => {
        const sPartKey = stock.partName.toString().trim().toLowerCase();
        let filteredTrans = transByPart.get(sPartKey) || [];
        if (filteredTrans.length === 0) {
          if (!partialMatchCache.has(sPartKey)) {
            const matchedKeys = uniqueTransPartNames.filter(tKey => 
              tKey.includes(sPartKey) || sPartKey.includes(tKey)
            );
            partialMatchCache.set(sPartKey, matchedKeys);
          }
          const matchedKeys = partialMatchCache.get(sPartKey)!;
          filteredTrans = matchedKeys.flatMap(k => transByPart.get(k) || []);
        }

        const onlyDayTrans = filteredTrans.filter(t => t.parsedDate && t.parsedDate >= startOfD && t.parsedDate <= endOfD);

        onlyDayTrans.forEach(t => {
          const type = t.type.toLowerCase();
          const qty = t.quantity;
          if (activeTab === 'bonding') {
            const type = t.type.toLowerCase();
            const qty = t.quantity;
            if (type.includes('metal') && type.includes('in')) dayTotals.metalStoreIn += qty;
            else if (type.includes('chemical') && type.includes('in')) dayTotals.chemicalStoreIn += qty;
            else if (type.includes('phosphate') && type.includes('in')) dayTotals.phosphateIn += qty;
            else if (type.includes('mold') && type.includes('in')) dayTotals.moldIn += qty;
            else if (type.includes('injc') && type.includes('out')) dayTotals.injcMoldOut += qty;
            else if (type.includes('oil') && type.includes('seal') && type.includes('out')) dayTotals.oilSealOut += qty;
            else if (type.includes('hvcm') && type.includes('out')) dayTotals.hvcmOut += qty;
            else if (type.includes('rejection') && type.includes('out')) dayTotals.rejectionOutToMetalStore += qty;
            else if (type.includes('rejection') && !type.includes('in')) dayTotals.rejectionOutToMetalStore += qty;
          } else if (activeTab === 'auto-clave') {
            const type = t.type.toLowerCase();
            const qty = t.quantity;
            if (type.includes('prod') && type.includes('in')) dayTotals.autoClaveProdIn += qty;
            else if (type.includes('mini') && type.includes('store') && type.includes('in')) dayTotals.autoClaveMiniStoreIn += qty;
            else if (type.includes('metal') && type.includes('in')) dayTotals.autoClaveMetalIn += qty;
            else if (type.includes('rework') && type.includes('in')) dayTotals.autoClaveReworkIn += qty;
            else if (type.includes('metal') && type.includes('out')) dayTotals.autoClaveMetalOut += qty;
            else if (type.includes('reject') && type.includes('rps')) dayTotals.autoClaveRejectionOut += qty;
          } else if (activeTab === 'phosphate') {
            const type = t.type.toLowerCase();
            const qty = t.quantity;
            if (type.includes('metal') && type.includes('in')) dayTotals.metalStoreIn += qty;
            else if (type.includes('chemical') && type.includes('in')) dayTotals.chemicalStoreIn += qty;
            else if (type.includes('bonding') && type.includes('out')) dayTotals.phosphateOutToBonding += qty;
            else if (type.includes('reject') && type.includes('rps')) dayTotals.rejectionOutToRps += qty;
          } else if (activeTab === 'oil-seal') {
            if (type.includes('mold') && type.includes('in')) dayTotals.moldIn += qty;
            else if (type.includes('re work in') || type.includes('rework in')) dayTotals.reworkIn += qty;
            else if (type.includes('qc out')) dayTotals.qcOut += qty;
            else if (type.includes('rejection to rps out') || type.includes('rejection')) dayTotals.rejectionOut += qty;
          } else if (activeTab === 'extrusion') {
            if ((type.includes('prod') || type.includes('production')) && type.includes('in')) { dayTotals.extrusionProdIn += qty; dayTotals.totalIn += qty; }
            else if (type.includes('re work in') || type.includes('rework in')) { dayTotals.reworkIn += qty; dayTotals.totalIn += qty; }
            else if (type.includes('metal') && type.includes('in')) { dayTotals.extrusionMetalIn += qty; dayTotals.totalIn += qty; }
            else if (type.includes('mini') && type.includes('store') && type.includes('in')) { dayTotals.extrusionMiniStoreIn += qty; dayTotals.totalIn += qty; }
            else if (type.includes('reject') && type.includes('rps')) { dayTotals.rejectionOutToRps += qty; dayTotals.totalOut += qty; }
            else if (type.includes('fg') && type.includes('out')) { dayTotals.fgOut += qty; dayTotals.totalOut += qty; }
            else if (type.includes('trim') && type.includes('out')) { dayTotals.extrusionTrimOut += qty; dayTotals.totalOut += qty; }
            else if (type.includes('qc') && type.includes('out')) { dayTotals.qcOut += qty; dayTotals.totalOut += qty; }
          } else if (activeTab === 'molding') {
            if (type.includes('rejection out to rps')) dayTotals.rejectionOutToRps += qty;
            else if (type.includes('oil seal trimming out')) dayTotals.oilSealTrimmingOut += qty;
            else if (type.includes('trimming out')) dayTotals.trimmingOut += qty;
          } else if (activeTab === 'quality') {
            const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
            const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

            if (type.includes('fg rework') && isRecd) dayTotals.fgReworkIn += qty;
            else if (type.includes('metal') && isRecd) dayTotals.metalStoreIn += qty;
            else if (type.includes('customer') && isRecd) dayTotals.customerRejectionIn += qty;
            else if (((type.includes('oil') && type.includes('seal')) || type.includes('bonding')) && isRecd) dayTotals.oilSealTrimmingIn += qty;
            else if (type.includes('trim') && isRecd) dayTotals.trimmingIn += qty;
            else if (type.includes('extru') && isRecd) dayTotals.extrusionIn += qty;
            else if (type.includes('reject') && (isSent || !isRecd)) dayTotals.rejectionOutToRps += qty;
            else if (type.includes('metal') && isSent) dayTotals.metalStoreOut += qty;
            else if (((type.includes('oil') && type.includes('seal')) || type.includes('bonding')) && isSent) dayTotals.oilSealTrimmingOut += qty;
            else if (type.includes('trim') && isSent) dayTotals.trimmingOut += qty;
            else if (type.includes('fg') && isSent) dayTotals.fgOut += qty;
            else if (type.includes('extru') && isSent) dayTotals.extrusionOut += qty;

            if (isRecd) dayTotals.totalIn += qty;
            else if (isSent || type.includes('reject')) dayTotals.totalOut += qty;
          } else if (activeTab === 'mini-store') {
            const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
            const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

            if (type.includes('compound') && isRecd) dayTotals.compoundIn += qty;
            else if (type.includes('mold return') && isRecd) dayTotals.moldReturnIn += qty;
            else if (type.includes('vendor') && isSent) dayTotals.vendorOut += qty;
            else if (type.includes('inject') && isSent) dayTotals.injectOut += qty;
            else if ((type.includes('oil seal') || type.includes('bonding')) && isSent) dayTotals.oilSealOut += qty;
            else if (type.includes('mold') && isSent) dayTotals.moldOut += qty;
            else if (type.includes('extru') && isSent) dayTotals.extrusionOut += qty;
            else if (type.includes('auto') && type.includes('clave') && isSent) dayTotals.autoClaveOut += qty;
            else if (type.includes('lab') && isSent) dayTotals.labOut += qty;
            else if (type.includes('reject') && (isSent || !isRecd)) dayTotals.rejectionOutToRps += qty;

            if (isRecd) dayTotals.totalIn += qty;
            else if (isSent || type.includes('reject')) dayTotals.totalOut += qty;
          } else if (activeTab === 'fg-store') {
            const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
            const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

            if (type.includes('customer') && isRecd) dayTotals.customerRejectionIn += qty;
            else if (type.includes('qc') && isRecd) dayTotals.qcIn += qty;
            else if (type.includes('rework') && isRecd) dayTotals.reworkIn += qty;
            else if (type.includes('auto') && type.includes('clave') && isRecd) dayTotals.autoClaveIn += qty;
            else if (type.includes('reject') && type.includes('rps') && isSent) dayTotals.rejectionOutToRps += qty;
            else if (type.includes('qc') && type.includes('rework') && isSent) dayTotals.qcReworkOut += qty;
            else if (type.includes('fg') && isSent) dayTotals.fgOut += qty;

            if (isRecd) dayTotals.totalIn += qty;
            else if (isSent) dayTotals.totalOut += qty;
          } else if (activeTab === 'trimming') {
            const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
            const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

            if (type.includes('vendor') && isRecd) dayTotals.trimmingVendorIn += qty;
            else if (type.includes('qc rework')) dayTotals.trimmingQcReworkIn += qty;
            else if (type.includes('mold') && isRecd) dayTotals.trimmingMoldingIn += qty;
            else if (type.includes('metal') && isRecd) dayTotals.trimmingMetalStoreIn += qty;
            else if (type.includes('extru') && isRecd) dayTotals.trimmingExtrusionIn += qty;
            else if (type.includes('qc') && isSent) dayTotals.trimmingQcOut += qty;
            else if (type.includes('vendor') && isSent) dayTotals.trimmingVendorOut += qty;
            else if (type.includes('reject') && (isSent || !isRecd)) dayTotals.trimmingRejectionOutToRps += qty;

            if (isRecd) dayTotals.totalIn += qty;
            else if (isSent || type.includes('reject')) dayTotals.totalOut += qty;
          }
        });
      });

      if (activeTab === 'bonding') {
        dayTotals.totalIn = dayTotals.metalStoreIn + dayTotals.chemicalStoreIn + dayTotals.phosphateIn + dayTotals.moldIn;
        dayTotals.totalOut = dayTotals.injcMoldOut + dayTotals.oilSealOut + dayTotals.hvcmOut + dayTotals.rejectionOutToMetalStore;
      } else if (activeTab === 'auto-clave') {
        dayTotals.totalIn = dayTotals.autoClaveProdIn + dayTotals.autoClaveMiniStoreIn + dayTotals.autoClaveMetalIn + dayTotals.autoClaveReworkIn;
        dayTotals.totalOut = dayTotals.autoClaveRejectionOut + dayTotals.autoClaveMetalOut;
      } else if (activeTab === 'phosphate') {
        dayTotals.totalIn = dayTotals.metalStoreIn + dayTotals.chemicalStoreIn;
        dayTotals.totalOut = dayTotals.phosphateOutToBonding + dayTotals.rejectionOutToRps;
      } else if (activeTab === 'oil-seal') {
        dayTotals.totalIn = dayTotals.moldIn + dayTotals.reworkIn;
        dayTotals.totalOut = dayTotals.qcOut + dayTotals.rejectionOut;
      } else if (activeTab === 'extrusion') {
        dayTotals.totalIn = dayTotals.reworkIn + dayTotals.extrusionProdIn + dayTotals.extrusionMetalIn + dayTotals.extrusionMiniStoreIn;
        dayTotals.totalOut = dayTotals.rejectionOutToRps + dayTotals.fgOut + dayTotals.extrusionTrimOut + dayTotals.qcOut;
      } else if (activeTab === 'molding') {
        dayTotals.totalIn = 0;
        dayTotals.totalOut = dayTotals.rejectionOutToRps + dayTotals.oilSealTrimmingOut + dayTotals.trimmingOut;
      } else if (activeTab === 'quality') {
        dayTotals.totalIn = dayTotals.fgReworkIn + dayTotals.metalStoreIn + dayTotals.customerRejectionIn + dayTotals.oilSealTrimmingIn + dayTotals.trimmingIn + dayTotals.extrusionIn;
        dayTotals.totalOut = dayTotals.rejectionOutToRps + dayTotals.metalStoreOut + dayTotals.oilSealTrimmingOut + dayTotals.trimmingOut + dayTotals.fgOut + dayTotals.extrusionOut;
      } else if (activeTab === 'mini-store') {
        dayTotals.totalIn = dayTotals.compoundIn + dayTotals.moldReturnIn;
        dayTotals.totalOut = dayTotals.vendorOut + dayTotals.injectOut + dayTotals.oilSealOut + dayTotals.moldOut + dayTotals.extrusionOut + dayTotals.autoClaveOut + dayTotals.labOut + dayTotals.rejectionOutToRps;
      } else if (activeTab === 'fg-store') {
        dayTotals.totalIn = dayTotals.customerRejectionIn + dayTotals.qcIn + dayTotals.autoClaveIn;
        dayTotals.totalOut = dayTotals.rejectionOutToRps + dayTotals.qcReworkOut + dayTotals.fgOut;
      } else if (activeTab === 'trimming') {
        dayTotals.totalIn = dayTotals.trimmingVendorIn + dayTotals.trimmingQcReworkIn + dayTotals.trimmingMoldingIn + dayTotals.trimmingMetalStoreIn + dayTotals.trimmingExtrusionIn;
        dayTotals.totalOut = dayTotals.trimmingQcOut + dayTotals.trimmingVendorOut + dayTotals.trimmingRejectionOutToRps;
      }

      return dayTotals;
    });
  }, [selectedMonth, openingStocks, transactions, activeTab]);

  const sortedDailySummary = useMemo(() => {
    if (!dailySummary.length) return [];
    if (!dailySortConfig) return dailySummary;

    return [...dailySummary].sort((a, b) => {
      const aVal = a[dailySortConfig.key];
      const bVal = b[dailySortConfig.key];

      if (aVal < bVal) return dailySortConfig.direction === 'asc' ? -1 : 1;
      if (aVal > bVal) return dailySortConfig.direction === 'asc' ? 1 : -1;
      return 0;
    });
  }, [dailySummary, dailySortConfig]);

  const dailyTotals = useMemo(() => {
    return dailySummary.reduce((acc, day) => {
      Object.keys(day).forEach(key => {
        if (key !== 'date') {
          acc[key] = (acc[key] || 0) + (day as any)[key];
        }
      });
      return acc;
    }, {} as any);
  }, [dailySummary]);

  const handleDailySort = (key: string) => {
    let direction: 'asc' | 'desc' = 'asc';
    if (dailySortConfig && dailySortConfig.key === key && dailySortConfig.direction === 'asc') {
      direction = 'desc';
    }
    setDailySortConfig({ key, direction });
  };

  const DailySortIcon = ({ field }: { field: string }) => {
    if (!dailySortConfig || dailySortConfig.key !== field) return <ArrowUpDown className="w-3 h-3 text-slate-300" />;
    return dailySortConfig.direction === 'asc' ? <ArrowUp className="w-3 h-3 text-blue-600" /> : <ArrowDown className="w-3 h-3 text-blue-600" />;
  };

  const summaryData = useMemo(() => {
    const filtered = allItemsWithActivity.filter(item => {
      // 0. Filter out items with empty names
      if (!item.itemId || item.itemId.trim().length === 0) return false;

      // 1. Filter out items with zero activity
      if (!item.hasActivity) return false;

      // 2. Excel-like Part Name filter
      if (selectedPartNames.length > 0) {
        return selectedPartNames.includes(item.itemId);
      }

      return true;
    });

    // Apply Sorting
    if (sortConfig.field) {
      filtered.sort((a, b) => {
        const aVal = a[sortConfig.field];
        const bVal = b[sortConfig.field];

        if (typeof aVal === 'string' && typeof bVal === 'string') {
          return sortConfig.direction === 'asc' 
            ? aVal.localeCompare(bVal) 
            : bVal.localeCompare(aVal);
        }

        if (typeof aVal === 'number' && typeof bVal === 'number') {
          return sortConfig.direction === 'asc' 
            ? aVal - bVal 
            : bVal - aVal;
        }

        return 0;
      });
    }

    return filtered;
  }, [allItemsWithActivity, selectedPartNames, sortConfig]);

  const columnTotals = useMemo(() => {
    return summaryData.reduce((acc, item) => ({
      openingStock: acc.openingStock + item.openingStock,
      vendorOpeningStock: acc.vendorOpeningStock + item.vendorOpeningStock,
      // Oil Seal
      moldIn: acc.moldIn + item.moldIn,
      reworkIn: acc.reworkIn + item.reworkIn,
      qcOut: acc.qcOut + item.qcOut,
      rejectionOut: acc.rejectionOut + item.rejectionOut,
      // Bonding
      chemicalStoreIn: acc.chemicalStoreIn + item.chemicalStoreIn,
      phosphateIn: acc.phosphateIn + item.phosphateIn,
      injcMoldOut: acc.injcMoldOut + item.injcMoldOut,
      hvcmOut: acc.hvcmOut + item.hvcmOut,
      rejectionOutToMetalStore: acc.rejectionOutToMetalStore + item.rejectionOutToMetalStore,
      // Phosphate
      phosphateOutToBonding: acc.phosphateOutToBonding + (item.phosphateOutToBonding || 0),
      // Auto Clave
      autoClaveProdIn: (acc.autoClaveProdIn || 0) + (item.autoClaveProdIn || 0),
      autoClaveMiniStoreIn: (acc.autoClaveMiniStoreIn || 0) + (item.autoClaveMiniStoreIn || 0),
      autoClaveMetalIn: (acc.autoClaveMetalIn || 0) + (item.autoClaveMetalIn || 0),
      autoClaveReworkIn: (acc.autoClaveReworkIn || 0) + (item.autoClaveReworkIn || 0),
      autoClaveRejectionOut: (acc.autoClaveRejectionOut || 0) + (item.autoClaveRejectionOut || 0),
      autoClaveMetalOut: (acc.autoClaveMetalOut || 0) + (item.autoClaveMetalOut || 0),
      // Quality
      fgReworkIn: acc.fgReworkIn + item.fgReworkIn,
      metalStoreIn: acc.metalStoreIn + item.metalStoreIn,
      customerRejectionIn: acc.customerRejectionIn + item.customerRejectionIn,
      oilSealTrimmingIn: acc.oilSealTrimmingIn + item.oilSealTrimmingIn,
      trimmingIn: acc.trimmingIn + item.trimmingIn,
      extrusionIn: acc.extrusionIn + (item.extrusionIn || 0),
      // Mini Store
      compoundIn: acc.compoundIn + item.compoundIn,
      moldReturnIn: acc.moldReturnIn + item.moldReturnIn,
      vendorOut: acc.vendorOut + item.vendorOut,
      injectOut: acc.injectOut + item.injectOut,
      oilSealOut: acc.oilSealOut + item.oilSealOut,
      moldOut: acc.moldOut + item.moldOut,
      extrusionOut: acc.extrusionOut + item.extrusionOut,
      autoClaveOut: acc.autoClaveOut + item.autoClaveOut,
      labOut: acc.labOut + item.labOut,
      // Common Out
      rejectionOutToRps: acc.rejectionOutToRps + item.rejectionOutToRps,
      metalStoreOut: acc.metalStoreOut + item.metalStoreOut,
      oilSealTrimmingOut: acc.oilSealTrimmingOut + item.oilSealTrimmingOut,
      trimmingOut: acc.trimmingOut + item.trimmingOut,
      fgOut: acc.fgOut + item.fgOut,
      // Trimming
      trimmingVendorIn: acc.trimmingVendorIn + item.trimmingVendorIn,
      trimmingQcReworkIn: acc.trimmingQcReworkIn + item.trimmingQcReworkIn,
      trimmingMoldingIn: acc.trimmingMoldingIn + item.trimmingMoldingIn,
      trimmingMetalStoreIn: acc.trimmingMetalStoreIn + item.trimmingMetalStoreIn,
      trimmingExtrusionIn: acc.trimmingExtrusionIn + item.trimmingExtrusionIn,
      trimmingQcOut: acc.trimmingQcOut + item.trimmingQcOut,
      trimmingVendorOut: acc.trimmingVendorOut + item.trimmingVendorOut,
      trimmingRejectionOutToRps: acc.trimmingRejectionOutToRps + item.trimmingRejectionOutToRps,
      // FG Store
      qcIn: acc.qcIn + (item.qcIn || 0),
      autoClaveIn: acc.autoClaveIn + (item.autoClaveIn || 0),
      qcReworkOut: acc.qcReworkOut + (item.qcReworkOut || 0),
      // Extrusion
      extrusionProdIn: acc.extrusionProdIn + (item.extrusionProdIn || 0),
      extrusionMetalIn: acc.extrusionMetalIn + (item.extrusionMetalIn || 0),
      extrusionMiniStoreIn: acc.extrusionMiniStoreIn + (item.extrusionMiniStoreIn || 0),
      extrusionTrimOut: acc.extrusionTrimOut + (item.extrusionTrimOut || 0),
      // Common
      totalIn: acc.totalIn + (item.totalIn || 0),
      totalOut: acc.totalOut + (item.totalOut || 0),
      currentStock: acc.currentStock + (item.currentStock || 0),
      vendorStock: acc.vendorStock + (item.vendorStock || 0),
      totalStock: acc.totalStock + (item.totalStock || 0),
      nextMonthOpeningStock: (acc.nextMonthOpeningStock || 0) + (item.nextMonthOpeningStock || 0),
    }), {
      openingStock: 0,
      moldIn: 0, reworkIn: 0, qcOut: 0, rejectionOut: 0,
      chemicalStoreIn: 0, phosphateIn: 0, injcMoldOut: 0, hvcmOut: 0, rejectionOutToMetalStore: 0,
      phosphateOutToBonding: 0,
      fgReworkIn: 0, metalStoreIn: 0, customerRejectionIn: 0, oilSealTrimmingIn: 0, trimmingIn: 0,
      compoundIn: 0, moldReturnIn: 0, vendorOut: 0, injectOut: 0, oilSealOut: 0, moldOut: 0, autoClaveOut: 0, labOut: 0,
      rejectionOutToRps: 0, metalStoreOut: 0, oilSealTrimmingOut: 0, trimmingOut: 0, fgOut: 0, extrusionOut: 0,
      autoClaveProdIn: 0, autoClaveMiniStoreIn: 0, autoClaveMetalIn: 0, autoClaveReworkIn: 0, autoClaveRejectionOut: 0, autoClaveMetalOut: 0,
      trimmingVendorIn: 0, trimmingQcReworkIn: 0, trimmingMoldingIn: 0, trimmingMetalStoreIn: 0, trimmingExtrusionIn: 0,
      trimmingQcOut: 0, trimmingVendorOut: 0, trimmingRejectionOutToRps: 0,
      qcIn: 0, autoClaveIn: 0, qcReworkOut: 0,
      extrusionProdIn: 0, extrusionMetalIn: 0, extrusionMiniStoreIn: 0, extrusionTrimOut: 0, extrusionIn: 0,
      totalIn: 0, totalOut: 0, currentStock: 0, nextMonthOpeningStock: 0, vendorOpeningStock: 0,
      vendorStock: 0, totalStock: 0
    });
  }, [summaryData]);

  const allTransactions = useMemo(() => {
    let trans = Object.values(dataCache).flatMap(cache => (cache as any).transactions as Transaction[]);
    
    const start = startDate ? parse(startDate, 'yyyy-MM-dd', new Date()) : null;
    const end = endDate ? parse(endDate, 'yyyy-MM-dd', new Date()) : null;

    if (start || end) {
      trans = trans.filter(t => {
        const tDate = t.parsedDate;
        if (!tDate) return false;
        if (start && tDate < start) return false;
        if (end && tDate > end) return false;
        return true;
      });
    }
    
    return trans;
  }, [dataCache, startDate, endDate]);

  const duplicateCount = useMemo(() => {
    if (activeTab !== 'job-tracking') return 0;
    const jobIdToPartNames = new Map<string, Set<string>>();
    
    allTransactions.forEach(t => {
      const jobId = t.trackingNumber?.trim();
      if (!jobId) return;
      if (!jobIdToPartNames.has(jobId)) {
        jobIdToPartNames.set(jobId, new Set());
      }
      jobIdToPartNames.get(jobId)!.add(t.partName || 'Unknown Part');
    });

    let count = 0;
    jobIdToPartNames.forEach((parts) => {
      if (parts.size > 1) count++;
    });
    return count;
  }, [allTransactions, activeTab]);

  interface JobGroup {
    jobId: string;
    partName: string;
    transactions: Transaction[];
    totals: Record<string, number>;
    sortedTotalsEntries: [string, number][];
    uniqueKey: string;
    isDuplicate: boolean;
    duplicatePartNames: string[];
  }

  const jobTrackingData = useMemo(() => {
    if (activeTab !== 'job-tracking') return [];
    
    // First pass: identify Job IDs associated with multiple Part Names
    const jobIdToPartNames = new Map<string, Set<string>>();
    allTransactions.forEach(t => {
      const jobId = t.trackingNumber?.trim();
      if (!jobId) return;
      if (!jobIdToPartNames.has(jobId)) {
        jobIdToPartNames.set(jobId, new Set());
      }
      jobIdToPartNames.get(jobId)!.add(t.partName || 'Unknown Part');
    });

    // Group transactions by trackingNumber and partName
    const grouped = allTransactions.reduce((acc, t) => {
      const jobId = t.trackingNumber?.trim();
      if (!jobId) return acc; // Skip transactions without a Job #
      
      const partName = t.partName || 'Unknown Part';
      const groupKey = `${jobId}|${partName}`;
      
      const partNamesForThisJob = jobIdToPartNames.get(jobId);
      const isDuplicate = (partNamesForThisJob?.size || 0) > 1;
      
      if (!acc[groupKey]) {
        acc[groupKey] = {
          jobId,
          partName,
          transactions: [],
          totals: {},
          sortedTotalsEntries: [],
          uniqueKey: groupKey,
          isDuplicate,
          duplicatePartNames: Array.from(partNamesForThisJob || [])
        };
      }
      acc[groupKey].transactions.push(t);
      
      // Aggregate by type with department prefix to avoid collisions and allow sorting
      const typeKey = `${t.department}|${t.type}`;
      acc[groupKey].totals[typeKey] = (acc[groupKey].totals[typeKey] || 0) + t.quantity;
      
      return acc;
    }, {} as Record<string, JobGroup>);

    const result: JobGroup[] = Object.values(grouped);
    
    // Define the strict production flow order
    const departmentOrder: Record<string, number> = {
      'molding': 1,
      'oil-seal': 2,
      'trimming': 3,
      'quality': 4,
      'fg-store': 5,
      'bonding': 6,
      'phosphate': 7,
      'auto-clave': 8,
      'extrusion': 9,
      'mini-store': 10
    };

    const getSortScore = (dept: string, type: string) => {
      const deptNormalized = dept.toLowerCase();
      const deptScore = departmentOrder[deptNormalized] ?? 99;
      const typeLower = type.toLowerCase();
      // IN transactions first, then OUT/Rejection transactions
      const isOut = typeLower.includes('out') || typeLower.includes('rejection');
      const typeScore = isOut ? 1 : 0;
      return deptScore * 10 + typeScore;
    };

    let filtered = result;

    if (showDuplicatesOnly) {
      filtered = filtered.filter(item => item.isDuplicate);
    }

    if (jobSearchTerm) {
      filtered = filtered.filter(item => 
        item.jobId.toLowerCase().includes(jobSearchTerm.toLowerCase()) ||
        item.partName.toLowerCase().includes(jobSearchTerm.toLowerCase())
      );
    }

    // Sort by jobId and process internal sorting
    return filtered.sort((a, b) => a.jobId.localeCompare(b.jobId)).map(job => {
      // Sort transactions by production flow
      const sortedTransactions = [...job.transactions].sort((a, b) => {
        const scoreA = getSortScore(a.department, a.type);
        const scoreB = getSortScore(b.department, b.type);
        if (scoreA !== scoreB) return scoreA - scoreB;
        // Within same scoring group, sort by date
        return new Date(a.date).getTime() - new Date(b.date).getTime();
      });

      // Pre-sort totals for the bubbles
      const sortedTotalsEntries = Object.entries(job.totals).sort((a, b) => {
        const [deptA, typeA] = a[0].split('|');
        const [deptB, typeB] = b[0].split('|');
        return getSortScore(deptA, typeA) - getSortScore(deptB, typeB);
      });

      return {
        ...job,
        transactions: sortedTransactions,
        sortedTotalsEntries
      };
    });
  }, [allTransactions, activeTab, jobSearchTerm, showDuplicatesOnly]);
  
  const visibleColumnCount = useMemo(() => {
    let count = 1; // Part No. & Name
    if (showJobColumn) count += 1;
    if (activeTab !== 'molding') count += 1; // Inhouse Opening Stock
    if (activeTab === 'trimming') count += 3; // Vendor Opening Stock, Vendor Stock, Total Stock
    count += 1; // Total IN
    count += 1; // Total OUT
    count += 1; // Current Stock
    if (hasAnyNextMonthStock) count += 1;

    const inCols = activeTab === 'bonding' ? ['metalStoreIn', 'chemicalStoreIn', 'phosphateIn', 'moldIn'] :
                   activeTab === 'auto-clave' ? ['autoClaveProdIn', 'autoClaveMiniStoreIn', 'autoClaveMetalIn', 'autoClaveReworkIn'] :
                   activeTab === 'phosphate' ? ['metalStoreIn', 'chemicalStoreIn'] :
                   activeTab === 'oil-seal' || activeTab === 'molding' ? ['moldIn', 'reworkIn'] : 
                   activeTab === 'extrusion' ? ['reworkIn', 'extrusionProdIn', 'extrusionMetalIn', 'extrusionMiniStoreIn'] :
                   activeTab === 'quality' ? ['fgReworkIn', 'metalStoreIn', 'customerRejectionIn', 'oilSealTrimmingIn', 'trimmingIn', 'extrusionIn'] :
                   activeTab === 'mini-store' ? ['compoundIn', 'moldReturnIn'] :
                   (activeTab === 'fg-store') ? ['customerRejectionIn', 'qcIn', 'reworkIn', 'autoClaveIn'] :
                   ['trimmingVendorIn', 'trimmingQcReworkIn', 'trimmingMoldingIn', 'trimmingMetalStoreIn', 'trimmingExtrusionIn'];
    
    const outCols = activeTab === 'bonding' ? ['injcMoldOut', 'oilSealOut', 'hvcmOut', 'rejectionOutToMetalStore'] :
                    activeTab === 'auto-clave' ? ['autoClaveRejectionOut', 'autoClaveMetalOut'] :
                    activeTab === 'phosphate' ? ['phosphateOutToBonding', 'rejectionOutToRps'] :
                    activeTab === 'oil-seal' || activeTab === 'molding' ? ['qcOut', 'rejectionOut'] :
                    activeTab === 'extrusion' ? ['rejectionOutToRps', 'fgOut', 'extrusionTrimOut', 'qcOut'] :
                    activeTab === 'quality' ? ['rejectionOutToRps', 'metalStoreOut', 'oilSealTrimmingOut', 'trimmingOut', 'fgOut', 'extrusionOut'] :
                    activeTab === 'mini-store' ? ['vendorOut', 'injectOut', 'oilSealOut', 'moldOut', 'extrusionOut', 'autoClaveOut', 'labOut', 'rejectionOutToRps'] :
                    (activeTab === 'fg-store') ? ['rejectionOutToRps', 'qcReworkOut', 'fgOut'] :
                    ['trimmingQcOut', 'trimmingVendorOut', 'trimmingRejectionOutToRps'];

    inCols.forEach(col => {
      if (isColVisible(columnTotals[col as keyof typeof columnTotals] || 0)) count++;
    });
    outCols.forEach(col => {
      if (isColVisible(columnTotals[col as keyof typeof columnTotals] || 0)) count++;
    });

    return count;
  }, [activeTab, hideZeroColumns, columnTotals, hasAnyNextMonthStock]);

  const topScrollRef = useRef<HTMLDivElement>(null);
  const tableScrollRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const topScroll = topScrollRef.current;
    const tableScroll = tableScrollRef.current;

    if (!topScroll || !tableScroll) return;

    const handleTopScroll = () => {
      if (tableScroll.scrollLeft !== topScroll.scrollLeft) {
        tableScroll.scrollLeft = topScroll.scrollLeft;
      }
    };

    const handleTableScroll = () => {
      if (topScroll.scrollLeft !== tableScroll.scrollLeft) {
        topScroll.scrollLeft = tableScroll.scrollLeft;
      }
    };

    topScroll.addEventListener('scroll', handleTopScroll);
    tableScroll.addEventListener('scroll', handleTableScroll);

    // Sync width
    const tableElement = tableScroll.querySelector('table');
    const topScrollContent = topScroll.querySelector('div');
    
    let resizeObserver: ResizeObserver | null = null;
    if (tableElement && topScrollContent) {
      resizeObserver = new ResizeObserver((entries) => {
        for (let entry of entries) {
          if (entry.target === tableElement) {
            topScrollContent.style.width = `${entry.contentRect.width}px`;
          }
        }
      });
      resizeObserver.observe(tableElement);
    }

    return () => {
      topScroll.removeEventListener('scroll', handleTopScroll);
      tableScroll.removeEventListener('scroll', handleTableScroll);
      if (resizeObserver) resizeObserver.disconnect();
    };
  }, [summaryData, activeTab]);

  const handleSort = (field: SortField) => {
    setSortConfig(prev => ({
      field,
      direction: prev.field === field && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };

  const SortIcon = ({ field }: { field: SortField }) => {
    if (sortConfig.field !== field) return <ArrowUpDown className="w-3 h-3 text-black stroke-[2.5]" />;
    return sortConfig.direction === 'asc' 
      ? <ArrowUp className="w-3 h-3 text-black stroke-[3]" /> 
      : <ArrowDown className="w-3 h-3 text-black stroke-[3]" />;
  };

  // Unique part names for the Excel filter - ONLY those with activity in current view
  const allPartNames = useMemo(() => {
    const uniqueIds = new Set(
      allItemsWithActivity
        .filter(item => item.hasActivity && item.itemId && item.itemId.trim().length > 0)
        .map(item => item.itemId)
    );
    return Array.from(uniqueIds).sort();
  }, [allItemsWithActivity]);

  // Filtered part names for the dropdown search
  const filteredPartNamesInDropdown = useMemo(() => {
    if (!filterSearch) return allPartNames;
    
    // Find item IDs that match either name or job #
    const matchingItemIds = allItemsWithActivity
      .filter(item => 
        item.itemId.toLowerCase().includes(filterSearch.toLowerCase()) ||
        (item.jobId && item.jobId.toLowerCase().includes(filterSearch.toLowerCase()))
      )
      .map(item => item.itemId);
      
    return allPartNames.filter(name => matchingItemIds.includes(name));
  }, [allPartNames, filterSearch, allItemsWithActivity]);

  const addDailySummarySheetToWorkbook = (
    workbook: XLSX.WorkBook, 
    tabTitle: string, 
    selectedMonth: string, 
    activeTab: string, 
    sortedDailySummary: any[], 
    dailyTotals: any, 
    initialOpening: number,
    initialVendorOpening: number = 0
  ) => {
    if (sortedDailySummary.length === 0) return;

    // 1. Sort by date ascending (A-Z)
    const sortedData = [...sortedDailySummary].sort((a, b) => {
      const aDate = a.date instanceof Date ? a.date : new Date(a.date);
      const bDate = b.date instanceof Date ? b.date : new Date(b.date);
      return aDate.getTime() - bDate.getTime();
    });

    const headers = (activeTab === 'bonding') ?
      ['DATE', 'OPENING STOCK', 'METAL STORE IN', 'CHEMICAL STORE IN', 'PHOSPHATE IN', 'MOLD IN', 'TOTAL IN', 'INJC MOLD OUT', 'OIL SEAL OUT', 'HVCM OUT', 'REJECTION OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'auto-clave' ?
      ['DATE', 'OPENING STOCK', 'PROD IN', 'MINI STORE IN', 'METAL IN', 'REWORK IN', 'TOTAL IN', 'REJECTION OUT', 'METAL OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'phosphate' ?
      ['DATE', 'OPENING STOCK', 'METAL STORE IN', 'CHEMICAL STORE IN', 'TOTAL IN', 'PHOSPHATE OUT TO BONDING', 'REJECTION OUT TO RPS', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'oil-seal' ? 
      ['DATE', 'OPENING STOCK', 'MOLD IN', 'REWORK IN', 'TOTAL IN', 'QC OUT', 'REJECTION OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'extrusion' ?
      ['DATE', 'OPENING STOCK', 'REWORK IN', 'PROD IN', 'METAL IN EXTRUS', 'MINI STORE IN EXTRUS', 'TOTAL IN', 'REJECTION OUT TO RPS', 'FG OUT', 'TRIM OUT', 'QC OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'molding' ?
      ['DATE', 'REJECTION OUT TO RPS', 'OIL SEAL TRIMMING OUT', 'TRIMMING OUT', 'TOTAL OUT'] :
      activeTab === 'quality' ?
      ['DATE', 'OPENING STOCK', 'FG REWORK IN', 'METAL STORE IN', 'CUSTOMER REJECTION IN', 'OIL SEAL TRIMMING IN', 'TRIMMING IN', 'EXTRUSION IN', 'TOTAL IN', 'REJECTION OUT TO RPS', 'METAL STORE OUT', 'OIL SEAL TRIMMING OUT', 'TRIMMING OUT', 'FG OUT', 'EXTRUSION OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'mini-store' ?
      ['DATE', 'OPENING STOCK', 'COMPOUND IN', 'MOLD RETURN IN', 'TOTAL IN', 'VENDOR OUT', 'INJECT OUT', 'OIL SEAL OUT', 'MOLD OUT', 'EXTRUSION OUT', 'AUTOCLAVE OUT', 'LAB OUT', 'REJECTION OUT TO RPS', 'TOTAL OUT', 'TOTAL STOCK'] :
      activeTab === 'fg-store' ?
      ['DATE', 'OPENING STOCK', 'CUSTOMER REJECTION IN', 'QC IN', 'REWORK IN', 'AUTO CLAVE IN', 'TOTAL IN', 'REJECTION OUT TO RPS', 'QC REWORK OUT', 'FG OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      ['DATE', 'INHOUSE OPENING STOCK', 'VENDOR OPENING STOCK', 'TRIMMING VENDOR IN', 'QC REWORK IN', 'MOLD IN', 'METAL STORE IN', 'EXTRUSION IN', 'TOTAL IN', 'QC OUT', 'VENDOR OUT', 'REJECTION OUT TO RPS', 'TOTAL OUT', 'IN HOUSE STOCK', 'VENDOR STOCK', 'TOTAL STOCK'];

    const groupSize = headers.length;

    // 2. Build data rows with formulas
    const data = sortedData.map((row, index) => {
      const rowNum = index + 4; // Title(1) + Totals(2) + Headers(3) + Data starts at 4
      const dateStr = format(row.date instanceof Date ? row.date : new Date(row.date), 'd-MMM-yyyy');
      
      // Opening Stock Formula: 
      // Row 4 (first data row) uses initialOpening
      // Row 5+ uses Current Stock of previous row (last column of previous row)
      const openingStock = index === 0 ? initialOpening : { f: `${XLSX.utils.encode_col(activeTab === 'trimming' ? 13 : groupSize - 1)}${rowNum - 1}` };
      const vendorOpeningStock = index === 0 ? initialVendorOpening : { f: `${XLSX.utils.encode_col(14)}${rowNum - 1}` };

      let rowData: any[] = [dateStr, openingStock];
      
      if (activeTab === 'bonding') {
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum}:${XLSX.utils.encode_col(5)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(7)}${rowNum}:${XLSX.utils.encode_col(10)}${rowNum})` };
        const currentStockFormula = { f: `${XLSX.utils.encode_col(6)}${rowNum}-${XLSX.utils.encode_col(11)}${rowNum}` };
        rowData.push(row.metalStoreIn, row.chemicalStoreIn, row.phosphateIn, row.moldIn, totalInFormula, row.injcMoldOut, row.oilSealOut, row.hvcmOut, row.rejectionOutToMetalStore, totalOutFormula, currentStockFormula);
      } else if (activeTab === 'auto-clave') {
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum}:${XLSX.utils.encode_col(5)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(7)}${rowNum}:${XLSX.utils.encode_col(8)}${rowNum})` };
        const currentStockFormula = { f: `${XLSX.utils.encode_col(6)}${rowNum}-${XLSX.utils.encode_col(9)}${rowNum}` };
        rowData.push(row.autoClaveProdIn, row.autoClaveMiniStoreIn, row.autoClaveMetalIn, row.autoClaveReworkIn, totalInFormula, row.autoClaveRejectionOut, row.autoClaveMetalOut, totalOutFormula, currentStockFormula);
      } else if (activeTab === 'phosphate') {
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum}:${XLSX.utils.encode_col(3)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(5)}${rowNum}:${XLSX.utils.encode_col(6)}${rowNum})` };
        const currentStockFormula = { f: `${XLSX.utils.encode_col(4)}${rowNum}-${XLSX.utils.encode_col(7)}${rowNum}` };
        rowData.push(row.metalStoreIn, row.chemicalStoreIn, totalInFormula, row.phosphateOutToBonding, row.rejectionOutToRps, totalOutFormula, currentStockFormula);
      } else if (activeTab === 'oil-seal') {
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum}:${XLSX.utils.encode_col(3)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(5)}${rowNum}:${XLSX.utils.encode_col(6)}${rowNum})` };
        const currentStockFormula = { f: `${XLSX.utils.encode_col(4)}${rowNum}-${XLSX.utils.encode_col(7)}${rowNum}` };
        rowData.push(row.moldIn, row.reworkIn, totalInFormula, row.qcOut, row.rejectionOut, totalOutFormula, currentStockFormula);
      } else if (activeTab === 'extrusion') {
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum}:${XLSX.utils.encode_col(5)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(7)}${rowNum}:${XLSX.utils.encode_col(10)}${rowNum})` };
        const currentStockFormula = { f: `${XLSX.utils.encode_col(6)}${rowNum}-${XLSX.utils.encode_col(11)}${rowNum}` };
        rowData.push(row.reworkIn, row.extrusionProdIn, row.extrusionMetalIn, row.extrusionMiniStoreIn, totalInFormula, row.rejectionOutToRps, row.fgOut, row.extrusionTrimOut, row.qcOut, totalOutFormula, currentStockFormula);
      } else if (activeTab === 'molding') {
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum}:${XLSX.utils.encode_col(3)}${rowNum})` };
        rowData = [dateStr, row.rejectionOutToRps, row.oilSealTrimmingOut, row.trimmingOut, totalOutFormula];
      } else if (activeTab === 'quality') {
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum}:${XLSX.utils.encode_col(7)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(9)}${rowNum}:${XLSX.utils.encode_col(14)}${rowNum})` };
        const currentStockFormula = { f: `${XLSX.utils.encode_col(8)}${rowNum}-${XLSX.utils.encode_col(15)}${rowNum}` };
        rowData.push(row.fgReworkIn, row.metalStoreIn, row.customerRejectionIn, row.oilSealTrimmingIn, row.trimmingIn, row.extrusionIn, totalInFormula, row.rejectionOutToRps, row.metalStoreOut, row.oilSealTrimmingOut, row.trimmingOut, row.fgOut, row.extrusionOut, totalOutFormula, currentStockFormula);
      } else if (activeTab === 'mini-store') {
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum}:${XLSX.utils.encode_col(3)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(5)}${rowNum}:${XLSX.utils.encode_col(12)}${rowNum})` };
        const currentStockFormula = { f: `${XLSX.utils.encode_col(4)}${rowNum}-${XLSX.utils.encode_col(13)}${rowNum}` };
        rowData.push(row.compoundIn, row.moldReturnIn, totalInFormula, row.vendorOut, row.injectOut, row.oilSealOut, row.moldOut, row.extrusionOut, row.autoClaveOut, row.labOut, row.rejectionOutToRps, totalOutFormula, currentStockFormula);
      } else if (activeTab === 'fg-store') {
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(2)}${rowNum}:${XLSX.utils.encode_col(5)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(7)}${rowNum}:${XLSX.utils.encode_col(9)}${rowNum})` };
        const currentStockFormula = { f: `${XLSX.utils.encode_col(6)}${rowNum}-${XLSX.utils.encode_col(10)}${rowNum}` };
        rowData.push(row.customerRejectionIn, row.qcIn, row.reworkIn, row.autoClaveIn, totalInFormula, row.rejectionOutToRps, row.qcReworkOut, row.fgOut, totalOutFormula, currentStockFormula);
      } else {
        // Trimming
        rowData.push(vendorOpeningStock);
        const totalInFormula = { f: `SUM(${XLSX.utils.encode_col(1)}${rowNum},${XLSX.utils.encode_col(3)}${rowNum}:${XLSX.utils.encode_col(7)}${rowNum})` };
        const totalOutFormula = { f: `SUM(${XLSX.utils.encode_col(9)}${rowNum}:${XLSX.utils.encode_col(11)}${rowNum})` };
        const inHouseStockFormula = { f: `${XLSX.utils.encode_col(8)}${rowNum}-${XLSX.utils.encode_col(12)}${rowNum}` };
        const vendorStockFormula = { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(10)}${rowNum}-${XLSX.utils.encode_col(3)}${rowNum}` };
        const totalStockFormula = { f: `${XLSX.utils.encode_col(13)}${rowNum}+${XLSX.utils.encode_col(14)}${rowNum}` };
        rowData.push(row.trimmingVendorIn, row.trimmingQcReworkIn, row.trimmingMoldingIn, row.trimmingMetalStoreIn, row.trimmingExtrusionIn, totalInFormula, row.trimmingQcOut, row.trimmingVendorOut, row.trimmingRejectionOutToRps, totalOutFormula, inHouseStockFormula, vendorStockFormula, totalStockFormula);
      }
      
      return rowData;
    });

    // 3. Add Totals Row (Row 2) with formulas summing the data rows
    const lastDataRow = data.length + 3;
    const totalsRow: any[] = ['TOTALS', initialOpening];
    if (activeTab === 'trimming') totalsRow.push(initialVendorOpening);

    for (let c = (activeTab === 'trimming' ? 3 : 2); c < groupSize; c++) {
      const colLetter = XLSX.utils.encode_col(c);
      if (c === groupSize - 1 || (activeTab === 'trimming' && (c === 13 || c === 14))) {
        // Last column or stock columns in trimming
        totalsRow.push({ f: `${colLetter}${lastDataRow}` });
      } else {
        totalsRow.push({ f: `SUM(${colLetter}4:${colLetter}${lastDataRow})` });
      }
    }

    const sheetData = [
      [`${tabTitle} Daily Transaction Summary - ${selectedMonth}`],
      totalsRow,
      headers,
      ...data
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
    
    // Styling
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!worksheet[addr]) continue;
        worksheet[addr].s = {
          border: { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } },
          alignment: { horizontal: C === 0 ? 'left' : 'center', vertical: 'center' }
        };
        if (R === 0) {
          worksheet[addr].s.font = { bold: true, size: 14 };
          worksheet[addr].s.fill = { fgColor: { rgb: 'E9ECEF' } };
        } else if (R === 1 || R === 2) {
          worksheet[addr].s.font = { bold: true };
          worksheet[addr].s.fill = { fgColor: { rgb: R === 1 ? 'F8F9FA' : 'F1F3F5' } };
          
          const totalInCol = activeTab === 'bonding' ? 6 : activeTab === 'phosphate' ? 4 : activeTab === 'auto-clave' ? 6 : activeTab === 'extrusion' ? 5 : activeTab === 'oil-seal' ? 4 : activeTab === 'molding' ? -1 : activeTab === 'trimming' ? 7 : activeTab === 'quality' ? 8 : activeTab === 'mini-store' ? 4 : (activeTab === 'fg-store' ? 5 : 8);
          const totalOutCol = activeTab === 'bonding' ? 11 : activeTab === 'phosphate' ? 7 : activeTab === 'auto-clave' ? 9 : activeTab === 'extrusion' ? 10 : activeTab === 'oil-seal' ? 7 : activeTab === 'molding' ? 4 : activeTab === 'trimming' ? 11 : activeTab === 'quality' ? 15 : activeTab === 'mini-store' ? 13 : (activeTab === 'fg-store' ? 10 : 15);
          const currentStockCol = activeTab === 'molding' ? -1 : headers.length - 1;
          
          if (C === 1) worksheet[addr].s.font.color = { rgb: 'E67E22' }; // Opening
          else if (C === totalInCol) worksheet[addr].s.font.color = { rgb: '008000' }; // Total IN
          else if (C === totalOutCol) worksheet[addr].s.font.color = { rgb: 'FF0000' }; // Total Out
          else if (C === currentStockCol) worksheet[addr].s.font.color = { rgb: '0000FF' }; // Current Stock
        }
      }
    }
    
    worksheet['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }];
    worksheet['!cols'] = [{ wch: 20 }, ...new Array(headers.length - 1).fill({ wch: 15 })];

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Daily Summary');
  };

  const handleExportExcel = () => {
    if (summaryData.length === 0) return;

    // Create a new workbook and worksheet
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([]); // Start with empty sheet
    
    const dataRowsCount = summaryData.length;
    const lastRow = dataRowsCount + 3; // Row 1 (Title) + Row 2 (Total) + Row 3 (Header) + Data rows

    // Determine Title
    const tabTitle = activeTab === 'molding' ? 'Molding' : 
                     activeTab === 'oil-seal' ? 'Oil Seal Trimming' :
                     (activeTab === 'bonding' || activeTab === 'phosphate' || activeTab === 'auto-clave') ? (activeTab === 'bonding' ? 'Bonding' : activeTab === 'phosphate' ? 'Phosphate' : 'Auto Clave') :
                     activeTab === 'extrusion' ? 'Extrusion' :
                     activeTab === 'quality' ? 'Quality' : 
                     activeTab === 'mini-store' ? 'Mini Store' :
                     activeTab === 'fg-store' ? 'FG Store' :
                     'Trimming';

    const reportType = showJobSummary ? 'JOB SUMMARY' : 'STOCK REPORT';
    
    // Dynamic Column Headers based on Job Summary mode
    const col1Header = showJobSummary ? 'Job #' : (activeTab === 'mini-store' ? 'ITEM ID' : (activeTab === 'quality' || activeTab === 'fg-store' ? 'Part No. & Name' : 'Item ID'));
    const col2Header = showJobSummary ? (activeTab === 'mini-store' ? 'ITEM ID' : (activeTab === 'quality' || activeTab === 'fg-store' ? 'Part No. & Name' : 'Item ID')) : 'Job #';

    let dateInfo = '';
    if (selectedDates.length > 0) {
      const sortedDates = [...selectedDates].sort((a, b) => a.getTime() - b.getTime());
      if (sortedDates.length === 1) {
        dateInfo = format(sortedDates[0], 'dd-MMMM-yyyy');
      } else {
        dateInfo = `${format(sortedDates[0], 'dd-MMM')} to ${format(sortedDates[sortedDates.length-1], 'dd-MMM-yyyy')}`;
      }
    } else if (startDate && endDate) {
      dateInfo = `${format(startDate, 'dd-MMM')} to ${format(endDate, 'dd-MMM-yyyy')}`;
    } else if (selectedMonth) {
      try {
        const parsedMonth = parse(selectedMonth, 'MMM-yy', new Date());
        dateInfo = format(parsedMonth, 'MMMM yyyy');
      } catch (e) {
        dateInfo = selectedMonth;
      }
    } else {
      dateInfo = `Till ${format(new Date(), 'dd-MMMM-yyyy')}`;
    }

    const title = `${tabTitle} ${reportType} - ${dateInfo}`;

    // Row 3: Headers
    const headers = (activeTab === 'bonding') ? [
      col1Header, 
      col2Header,
      'TOTAL OPENING STOCK', 
      'METAL STORE IN',
      'CHEMICAL STORE IN',
      'PHOSPHATE IN',
      'MOLD IN',
      'TOTAL IN',
      'INJC MOLD OUT',
      'OIL SEAL OUT',
      'HVCM OUT',
      'REJECTION OUT',
      'TOTAL OUT',
      'CURRENT STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ] : activeTab === 'auto-clave' ? [
      col1Header, 
      col2Header,
      'TOTAL OPENING STOCK', 
      'PROD IN',
      'MINI STORE IN',
      'METAL IN',
      'REWORK IN',
      'TOTAL IN',
      'REJECTION OUT',
      'METAL OUT',
      'TOTAL OUT',
      'CURRENT STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ] : activeTab === 'phosphate' ? [
      col1Header, 
      col2Header,
      'TOTAL OPENING STOCK', 
      'METAL STORE IN',
      'CHEMICAL STORE IN',
      'TOTAL IN',
      'PHOSPHATE OUT TO BONDING',
      'REJECTION OUT TO RPS',
      'TOTAL OUT',
      'CURRENT STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ] : activeTab === 'oil-seal' ? [
      col1Header, 
      col2Header,
      'TOTAL OPENING STOCK', 
      'MOLD IN', 
      'REWORK IN', 
      'TOTAL IN', 
      'QC OUT', 
      'REJECTION OUT', 
      'TOTAL OUT', 
      'CURRENT STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ] : activeTab === 'extrusion' ? [
      col1Header, 
      col2Header,
      'TOTAL OPENING STOCK', 
      'REWORK IN', 
      'PROD IN',
      'METAL IN EXTRUS',
      'MINI STORE IN EXTRUS',
      'TOTAL IN', 
      'REJECTION OUT TO RPS',
      'FG OUT',
      'TRIM OUT',
      'QC OUT',
      'TOTAL OUT', 
      'CURRENT STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ] : activeTab === 'molding' ? [
      col1Header,
      col2Header,
      'REJECTION OUT TO RPS',
      'OIL SEAL TRIMMING OUT',
      'TRIMMING OUT',
      'TOTAL OUT'
    ] : activeTab === 'quality' ? [
      col1Header,
      col2Header,
      'TOTAL OPENING STOCK',
      'FG REWORK IN',
      'METAL STORE IN',
      'CUSTOMER REJECTION IN',
      'OIL SEAL TRIMMING IN',
      'TRIMMING IN',
      'EXTRUSION IN',
      'TOTAL IN',
      'REJECTION OUT TO RPS',
      'METAL STORE OUT',
      'OIL SEAL TRIMMING OUT',
      'TRIMMING OUT',
      'FG OUT',
      'EXTRUSION OUT',
      'TOTAL OUT',
      'CURRENT STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ] : activeTab === 'mini-store' ? [
      col1Header,
      col2Header,
      'OPENING STOCK',
      'COMPOUND IN',
      'MOLD RETURN IN',
      'TOTAL IN',
      'VENDOR OUT',
      'INJECT OUT',
      'OIL SEAL OUT',
      'MOLD OUT',
      'EXTRUSION OUT',
      'AUTOCLAVE OUT',
      'LAB OUT',
      'REJECTION OUT TO RPS',
      'TOTAL OUT',
      'TOTAL STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ] : activeTab === 'fg-store' ? [
      col1Header,
      col2Header,
      'TOTAL OPENING STOCK',
      'CUSTOMER REJECTION IN',
      'QC IN',
      'REWORK IN',
      'AUTO CLAVE IN',
      'TOTAL IN',
      'REJECTION OUT TO RPS',
      'QC REWORK OUT',
      'FG OUT',
      'TOTAL OUT',
      'CURRENT STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ] : [
      // Trimming
      col1Header,
      col2Header,
      'INHOUSE OPENING STOCK',
      'VENDOR OPENING STOCK',
      'TRIMMING VENDOR IN',
      'QC REWORK IN',
      'MOLD IN',
      'METAL STORE IN',
      'EXTRUSION IN',
      'TOTAL IN',
      'QC OUT',
      'VENDOR OUT',
      'REJECTION OUT TO RPS',
      'TOTAL OUT',
      'IN HOUSE STOCK',
      'VENDOR STOCK',
      'TOTAL STOCK',
      ...(hasAnyNextMonthStock ? ['PHYSICAL STOCK'] : [])
    ];

    // Row 1: TITLE
    const titleRow = new Array(headers.length).fill('');
    titleRow[0] = title;
    XLSX.utils.sheet_add_aoa(worksheet, [titleRow], { origin: 'A1' });
    
    // Merge title row
    worksheet['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }
    ];

    // Row 2: TOTALS (with formulas summing the data rows)
    const totalsRow = headers.map((h, i) => {
      if (i === 0 || i === 1) return i === 0 ? 'Total' : '';
      const colLetter = XLSX.utils.encode_col(i);
      return { f: `SUM(${colLetter}4:${colLetter}${lastRow})` };
    });
    XLSX.utils.sheet_add_aoa(worksheet, [totalsRow], { origin: 'A2' });
    XLSX.utils.sheet_add_aoa(worksheet, [headers], { origin: 'A3' });

    // Row 4+: Data with formulas for calculated columns
    const dataRows = summaryData.map((s, index) => {
      const rowNum = index + 4; // Data starts at Row 4
      const col1Val = showJobSummary ? s.jobId : s.itemId;
      const col2Val = showJobSummary ? s.partName : s.jobId;

      if (activeTab === 'bonding') {
        return [
          col1Val,
          col2Val,
          s.openingStock,
          s.metalStoreIn,
          s.chemicalStoreIn,
          s.phosphateIn,
          s.moldIn,
          { f: `C${rowNum}+D${rowNum}+E${rowNum}+F${rowNum}+G${rowNum}` }, // Total IN
          s.injcMoldOut,
          s.oilSealOut,
          s.hvcmOut,
          s.rejectionOutToMetalStore,
          { f: `I${rowNum}+J${rowNum}+K${rowNum}+L${rowNum}` }, // Total OUT
          { f: `H${rowNum}-M${rowNum}` }, // Current Stock
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      } else if (activeTab === 'auto-clave') {
        return [
          col1Val,
          col2Val,
          s.openingStock,
          s.autoClaveProdIn,
          s.autoClaveMiniStoreIn,
          s.autoClaveMetalIn,
          s.autoClaveReworkIn,
          { f: `C${rowNum}+D${rowNum}+E${rowNum}+F${rowNum}+G${rowNum}` }, // Total IN
          s.autoClaveRejectionOut,
          s.autoClaveMetalOut,
          { f: `I${rowNum}+J${rowNum}` }, // Total OUT
          { f: `H${rowNum}-K${rowNum}` }, // Current Stock
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      } else if (activeTab === 'phosphate') {
        return [
          col1Val,
          col2Val,
          s.openingStock,
          s.metalStoreIn,
          s.chemicalStoreIn,
          { f: `C${rowNum}+D${rowNum}+E${rowNum}` }, // Total IN
          s.phosphateOutToBonding,
          s.rejectionOutToRps,
          { f: `G${rowNum}+H${rowNum}` }, // Total OUT
          { f: `F${rowNum}-I${rowNum}` }, // Current Stock
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      } else if (activeTab === 'oil-seal') {
        return [
          col1Val,
          col2Val,
          s.openingStock,
          s.moldIn,
          s.reworkIn,
          { f: `C${rowNum}+D${rowNum}+E${rowNum}` }, // Total IN
          s.qcOut,
          s.rejectionOut,
          { f: `G${rowNum}+H${rowNum}` }, // Total OUT
          { f: `F${rowNum}-I${rowNum}` }, // Current Stock
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      } else if (activeTab === 'extrusion') {
        return [
          col1Val,
          col2Val,
          s.openingStock,
          s.reworkIn,
          s.extrusionProdIn,
          s.extrusionMetalIn,
          s.extrusionMiniStoreIn,
          { f: `C${rowNum}+D${rowNum}+E${rowNum}+F${rowNum}+G${rowNum}` }, // Total IN
          s.rejectionOutToRps,
          s.fgOut,
          s.extrusionTrimOut,
          s.qcOut,
          { f: `I${rowNum}+J${rowNum}+K${rowNum}+L${rowNum}` }, // Total OUT
          { f: `H${rowNum}-M${rowNum}` }, // Current Stock
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      } else if (activeTab === 'molding') {
        return [
          col1Val,
          col2Val,
          s.rejectionOutToRps,
          s.oilSealTrimmingOut,
          s.trimmingOut,
          { f: `C${rowNum}+D${rowNum}+E${rowNum}` } // Total OUT
        ];
      } else if (activeTab === 'quality') {
        // Quality
        return [
          col1Header === 'Job #' ? s.jobId : s.itemId,
          col1Header === 'Job #' ? s.itemId : s.jobId,
          s.openingStock,
          s.fgReworkIn,
          s.metalStoreIn,
          s.customerRejectionIn,
          s.oilSealTrimmingIn,
          s.trimmingIn,
          s.extrusionIn,
          { f: `C${rowNum}+D${rowNum}+E${rowNum}+F${rowNum}+G${rowNum}+H${rowNum}+I${rowNum}` }, // Total IN
          s.rejectionOutToRps,
          s.metalStoreOut,
          s.oilSealTrimmingOut,
          s.trimmingOut,
          s.fgOut,
          s.extrusionOut,
          { f: `K${rowNum}+L${rowNum}+M${rowNum}+N${rowNum}+O${rowNum}+P${rowNum}` }, // Total OUT
          { f: `J${rowNum}-Q${rowNum}` }, // Current Stock
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      } else if (activeTab === 'mini-store') {
        // Mini Store
        return [
          col1Header === 'Job #' ? s.jobId : s.itemId,
          col1Header === 'Job #' ? s.itemId : s.jobId,
          s.openingStock,
          s.compoundIn,
          s.moldReturnIn,
          { f: `D${rowNum}+E${rowNum}` }, // Total IN (Compound + Mold Return)
          s.vendorOut,
          s.injectOut,
          s.oilSealOut,
          s.moldOut,
          s.extrusionOut,
          s.autoClaveOut,
          s.labOut,
          s.rejectionOutToRps,
          { f: `G${rowNum}+H${rowNum}+I${rowNum}+J${rowNum}+K${rowNum}+L${rowNum}+M${rowNum}+N${rowNum}` }, // Total OUT (G:N)
          { f: `C${rowNum}+F${rowNum}-O${rowNum}` }, // Total Stock (Opening + Total In - Total Out)
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      } else if (activeTab === 'fg-store') {
        return [
          col1Val,
          col2Val,
          s.openingStock,
          s.customerRejectionIn,
          s.qcIn,
          s.reworkIn,
          s.autoClaveIn,
          { f: `D${rowNum}+E${rowNum}+F${rowNum}+G${rowNum}` }, // Total IN (D:G)
          s.rejectionOutToRps,
          s.qcReworkOut,
          s.fgOut,
          { f: `I${rowNum}+J${rowNum}+K${rowNum}` }, // Total OUT (I:K)
          { f: `C${rowNum}+H${rowNum}-L${rowNum}` }, // Current Stock (Opening + Total In - Total Out)
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      } else {
        // Trimming
        return [
          col1Val,
          col2Val,
          s.openingStock,
          s.vendorOpeningStock,
          s.trimmingVendorIn,
          s.trimmingQcReworkIn,
          s.trimmingMoldingIn,
          s.trimmingMetalStoreIn,
          s.trimmingExtrusionIn,
          { f: `C${rowNum}+E${rowNum}+F${rowNum}+G${rowNum}+H${rowNum}+I${rowNum}` }, // Total IN
          s.trimmingQcOut,
          s.trimmingVendorOut,
          s.trimmingRejectionOutToRps,
          { f: `K${rowNum}+L${rowNum}+M${rowNum}` }, // Total OUT
          { f: `J${rowNum}-N${rowNum}` }, // IN HOUSE STOCK
          { f: `D${rowNum}+L${rowNum}-E${rowNum}` }, // VENDOR STOCK
          { f: `O${rowNum}+P${rowNum}` }, // TOTAL STOCK
          ...(hasAnyNextMonthStock ? [s.nextMonthOpeningStock] : [])
        ];
      }
    });
    XLSX.utils.sheet_add_aoa(worksheet, dataRows, { origin: 'A4' });

    // Apply styles (borders and bold headers)
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        if (!worksheet[cellAddress]) continue;
        
        // Initialize style object if it doesn't exist
        worksheet[cellAddress].s = {
          border: {
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
          },
          alignment: {
            horizontal: C === 0 ? 'left' : 'center',
            vertical: 'center'
          }
        };

        // Title styling (Row 1)
        if (R === 0) {
          worksheet[cellAddress].s.font = { bold: true, size: 14 };
          worksheet[cellAddress].s.alignment.horizontal = 'center';
          worksheet[cellAddress].s.fill = { fgColor: { rgb: 'E9ECEF' } };
        }

        // Bold for Row 2 (Totals) and Row 3 (Headers)
        if (R === 1 || R === 2) {
          worksheet[cellAddress].s.font = { bold: true };
          
          // Apply specific colors to important columns
          let totalInCol = -1;
          let totalOutCol = -1;
          let currentStockCol = -1;

          if (activeTab === 'oil-seal' || activeTab === 'bonding' || activeTab === 'phosphate' || activeTab === 'auto-clave' || activeTab === 'molding') {
            totalInCol = activeTab === 'phosphate' ? 5 : activeTab === 'bonding' ? 7 : activeTab === 'auto-clave' ? 7 : 5;
            totalOutCol = activeTab === 'phosphate' ? 8 : activeTab === 'bonding' ? 12 : activeTab === 'auto-clave' ? 10 : 8;
            currentStockCol = activeTab === 'phosphate' ? 9 : activeTab === 'bonding' ? 13 : activeTab === 'auto-clave' ? 11 : 9;
            if (activeTab === 'molding') {
              totalInCol = -1;
              totalOutCol = 5;
              currentStockCol = -1;
            }
          } else if (activeTab === 'extrusion') {
            totalInCol = 7;
            totalOutCol = 12;
            currentStockCol = 13;
          } else if (activeTab === 'trimming') {
            totalInCol = 9;
            totalOutCol = 13;
            currentStockCol = 14;
          } else if (activeTab === 'mini-store') {
            totalInCol = 5;
            totalOutCol = 14;
            currentStockCol = 15;
          } else if (activeTab === 'fg-store') {
            totalInCol = 7;
            totalOutCol = 11;
            currentStockCol = 12;
          } else {
            // Quality
            totalInCol = 9;
            totalOutCol = 16;
            currentStockCol = 17;
          }
          
          const nextMonthCol = currentStockCol + (activeTab === 'trimming' ? 3 : 1);

          if (C === 2) { // Total Opening Stock (Shifted to index 2)
            worksheet[cellAddress].s.font.color = { rgb: 'E67E22' }; // Dark Orange
          } else if (activeTab === 'trimming' && C === 3) { // Vendor Opening Stock
            worksheet[cellAddress].s.font.color = { rgb: 'E67E22' }; // Dark Orange
          } else if (C === totalInCol) { // Total IN
            worksheet[cellAddress].s.font.color = { rgb: '008000' }; // Green
          } else if (C === totalOutCol) { // Total Out
            worksheet[cellAddress].s.font.color = { rgb: 'FF0000' }; // Red
          } else if (C === currentStockCol) { // Current Stock / IN HOUSE STOCK
            worksheet[cellAddress].s.font.color = { rgb: '0000FF' }; // Blue
          } else if (activeTab === 'trimming' && C === 15) { // VENDOR STOCK
            worksheet[cellAddress].s.font.color = { rgb: '0000FF' }; // Blue
          } else if (activeTab === 'trimming' && C === 16) { // TOTAL STOCK
            worksheet[cellAddress].s.font.color = { rgb: '0000FF' }; // Blue
          } else if (hasAnyNextMonthStock && C === nextMonthCol) { // Next Month Opening
            worksheet[cellAddress].s.font.color = { rgb: '4B0082' }; // Indigo
          }

          if (R === 1) {
            worksheet[cellAddress].s.fill = { fgColor: { rgb: 'F8F9FA' } }; // Very light gray for totals
          } else {
            worksheet[cellAddress].s.fill = { fgColor: { rgb: 'F1F3F5' } }; // Light gray for headers
          }
        }
      }
    }

    // Set column widths for better readability
    if (activeTab === 'oil-seal') {
      worksheet['!cols'] = [
        { wch: 40 }, // Item ID
        { wch: 15 }, // Job #
        { wch: 18 }, // Total Opening Stock
        { wch: 15 }, // Mold IN
        { wch: 15 }, // Re Work IN
        { wch: 12 }, // Total IN
        { wch: 15 }, // QC Out
        { wch: 15 }, // Rejection OUT
        { wch: 12 }, // Total Out
        { wch: 12 }  // Current Stock
      ];
    } else if (activeTab === 'extrusion') {
      worksheet['!cols'] = [
        { wch: 40 }, // Part No. & Name
        { wch: 15 }, // Job #
        { wch: 18 }, // Total Opening Stock
        { wch: 15 }, // Rework IN
        { wch: 15 }, // Prod IN
        { wch: 15 }, // Metal IN Extrus
        { wch: 15 }, // Mini Store IN Extrus
        { wch: 12 }, // Total In
        { wch: 20 }, // Rejection OUT to RPS
        { wch: 12 }, // FG Out
        { wch: 12 }, // TRIM OUT
        { wch: 12 }, // QC Out
        { wch: 12 }, // Total Out
        { wch: 15 }  // Current Stock
      ];
    } else if (activeTab === 'bonding' || activeTab === 'phosphate' || activeTab === 'auto-clave') {
      worksheet['!cols'] = [
        { wch: 40 }, // Item ID
        { wch: 15 }, // Job #
        { wch: 18 }, // Total Opening Stock
        { wch: 15 }, // Metal Store IN
        { wch: 15 }, // Chemical Store IN
        { wch: 15 }, // Phosphate IN
        { wch: 15 }, // Mold IN
        { wch: 12 }, // Total IN
        { wch: 15 }, // Injc Mold Out
        { wch: 15 }, // Oil Seal Out
        { wch: 15 }, // Hvcm Out
        { wch: 15 }, // Rejection OUT
        { wch: 12 }, // Total Out
        { wch: 12 }  // Current Stock
      ];
    } else if (activeTab === 'trimming') {
      worksheet['!cols'] = [
        { wch: 50 }, // Item ID
        { wch: 12 }, // Job #
        { wch: 18 }, // Inhouse Opening Stock
        { wch: 18 }, // Vendor Opening Stock
        { wch: 15 }, // Trimming Vendor In
        { wch: 15 }, // QC ReWork IN
        { wch: 15 }, // Mold IN
        { wch: 15 }, // Metal Store IN
        { wch: 15 }, // Extrusion IN
        { wch: 12 }, // Total IN
        { wch: 15 }, // QC Out
        { wch: 15 }, // Vendor Out
        { wch: 15 }, // Rejection Out to RPS
        { wch: 12 }, // Total Out
        { wch: 15 }, // IN HOUSE STOCK
        { wch: 15 }, // VENDOR STOCK
        { wch: 15 }  // TOTAL STOCK
      ];
    } else {
      // Quality
      worksheet['!cols'] = [
        { wch: 50 }, // Part No. & Name
        { wch: 18 }, // Total Opening Stock
        { wch: 15 }, // FG Rework IN
        { wch: 15 }, // Metal Store IN
        { wch: 15 }, // Customer Rejection IN
        { wch: 15 }, // Oil Seal Trimming IN
        { wch: 15 }, // Trimming IN
        { wch: 15 }, // Extrusion IN
        { wch: 12 }, // Total IN
        { wch: 15 }, // Rejection OUT to RPS
        { wch: 15 }, // Metal Store OUT
        { wch: 15 }, // Oil Seal Trimming OUT
        { wch: 15 }, // Trimming OUT
        { wch: 15 }, // FG Out
        { wch: 15 }, // Extrusion OUT
        { wch: 12 }, // Total Out
        { wch: 12 }  // Current Stock
      ];
    }

    XLSX.utils.book_append_sheet(workbook, worksheet, showJobSummary ? 'Job Summary' : tabTitle);
    
    // Add Daily Summary Sheet
    const initialOpening = summaryData.reduce((sum, item) => sum + item.openingStock, 0);
    const initialVendorOpening = (activeTab === 'trimming' || activeTab === 'extrusion') ? summaryData.reduce((sum, item) => sum + (item.vendorOpeningStock || 0), 0) : 0;
    addDailySummarySheetToWorkbook(workbook, tabTitle, selectedMonth, activeTab, sortedDailySummary, dailyTotals, initialOpening, initialVendorOpening);

    // Generate filename with current date/month
    const fileName = `${tabTitle.replace(/\s+/g, '_')}_${reportType.replace(/\s+/g, '_')}_${dateInfo.replace(/\s+/g, '_')}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  const handleExportJobTracking = () => {
    if (jobTrackingData.length === 0) return;

    const wb = XLSX.utils.book_new();
    
    // Prepare data for Excel
    const excelData = jobTrackingData.map(job => {
      const row: any = {
        "Job #": job.jobId,
        "Part Name": job.partName
      };
      
      // Add totals
      Object.entries(job.totals as Record<string, number>).forEach(([type, qty]) => {
        row[type] = qty;
      });
      
      return row;
    });

    const ws = XLSX.utils.json_to_sheet(excelData);
    XLSX.utils.book_append_sheet(wb, ws, "Job Tracking Summary");

    // Also add detailed history sheet
    const detailedData = jobTrackingData.flatMap(job => 
      job.transactions.map((t: any) => ({
        "Job #": job.jobId,
        "Part Name": job.partName,
        "Department": t.department,
        "Date": t.date,
        "Type": t.type,
        "Quantity": t.quantity,
        "Shift": t.shift
      }))
    );

    const wsDetailed = XLSX.utils.json_to_sheet(detailedData);
    XLSX.utils.book_append_sheet(wb, wsDetailed, "Detailed History");

    XLSX.writeFile(wb, `Job_Tracking_${format(new Date(), 'yyyy-MM-dd')}.xlsx`);
  };

  const handleExportMonthlyDailyReport = () => {
    if (!selectedMonth || summaryData.length === 0) {
      alert("Please select a month first.");
      return;
    }

    let parsedMonth: Date;
    try {
      parsedMonth = parse(selectedMonth, 'MMM-yy', new Date());
      if (!isValid(parsedMonth)) {
        parsedMonth = parse(selectedMonth, 'MMM yyyy', new Date());
      }
    } catch (e) {
      alert("Could not parse the selected month.");
      return;
    }

    if (!isValid(parsedMonth)) {
      alert("Invalid month selection.");
      return;
    }

    const startOfM = startOfMonth(parsedMonth);
    const endOfM = endOfMonth(parsedMonth);
    const allDaysInMonth = eachDayOfInterval({ start: startOfM, end: endOfM });

    // 1. Identify days with transactions in the selected month
    const daysWithTransactions = new Set<string>();
    transactions.forEach(t => {
      if (t.parsedDate) {
        const isSameMonth = t.parsedDate.getMonth() === parsedMonth.getMonth() && 
                           t.parsedDate.getFullYear() === parsedMonth.getFullYear();
        if (isSameMonth) {
          daysWithTransactions.add(format(t.parsedDate, 'yyyy-MM-dd'));
        }
      }
    });

    const sortedActiveDays = allDaysInMonth.filter(day => 
      daysWithTransactions.has(format(day, 'yyyy-MM-dd'))
    );

    if (sortedActiveDays.length === 0) {
      alert("No transactions found for the selected month.");
      return;
    }

    const workbook = XLSX.utils.book_new();
    const vendorStocksByPart = new Map<string, number>();
    if (activeTab === 'trimming') {
      vendorOpeningStocks.forEach(vs => {
        const key = vs.partName.toString().trim().toLowerCase();
        const val = vs.monthlyStocks[selectedMonth] || 0;
        vendorStocksByPart.set(key, val);
      });
    }
     const tabTitle = activeTab === 'molding' ? 'Molding' : 
                      activeTab === 'oil-seal' || activeTab === 'bonding' || activeTab === 'phosphate' || activeTab === 'auto-clave' || activeTab === 'extrusion' ? 
                      (activeTab === 'oil-seal' ? 'Oil Seal Trimming' : activeTab === 'bonding' ? 'Bonding' : activeTab === 'phosphate' ? 'Phosphate' : activeTab === 'auto-clave' ? 'Auto Clave' : 'Extrusion') : 
                      activeTab === 'quality' ? 'Quality' : 
                      activeTab === 'mini-store' ? 'Mini Store' :
                      activeTab === 'fg-store' ? 'FG Store' :
                      'Trimming';

    // 2. Define Group Headers based on Tab
    const groupHeaders = (activeTab === 'bonding') ?
      ['OPENING STOCK', 'METAL STORE IN', 'CHEMICAL STORE IN', 'PHOSPHATE IN', 'MOLD IN', 'TOTAL IN', 'INJC MOLD OUT', 'OIL SEAL OUT', 'HVCM OUT', 'REJECTION OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'auto-clave' ?
      ['OPENING STOCK', 'PROD IN', 'MINI STORE IN', 'METAL IN', 'REWORK IN', 'TOTAL IN', 'REJECTION OUT', 'METAL OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'phosphate' ?
      ['OPENING STOCK', 'METAL STORE IN', 'CHEMICAL STORE IN', 'TOTAL IN', 'PHOSPHATE OUT TO BONDING', 'REJECTION OUT TO RPS', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'oil-seal' || activeTab === 'extrusion' ? 
      ['OPENING STOCK', 'MOLD IN', 'REWORK IN', 'TOTAL IN', 'QC OUT', 'REJECTION OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'molding' ?
      ['REJECTION OUT TO RPS', 'OIL SEAL TRIMMING OUT', 'TRIMMING OUT', 'TOTAL OUT'] :
      activeTab === 'quality' ?
      ['OPENING STOCK', 'FG REWORK IN', 'METAL STORE IN', 'CUSTOMER REJECTION IN', 'OIL SEAL TRIMMING IN', 'TRIMMING IN', 'EXTRUSION IN', 'TOTAL IN', 'REJECTION OUT TO RPS', 'METAL STORE OUT', 'OIL SEAL TRIMMING OUT', 'TRIMMING OUT', 'FG OUT', 'EXTRUSION OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      activeTab === 'mini-store' ?
      ['OPENING STOCK', 'COMPOUND IN', 'MOLD RETURN IN', 'TOTAL IN', 'VENDOR OUT', 'INJECT OUT', 'OIL SEAL OUT', 'MOLD OUT', 'EXTRUSION OUT', 'AUTOCLAVE OUT', 'LAB OUT', 'REJECTION OUT TO RPS', 'TOTAL OUT', 'TOTAL STOCK'] :
      activeTab === 'fg-store' ?
      ['OPENING STOCK', 'CUSTOMER REJECTION IN', 'QC IN', 'REWORK IN', 'AUTO CLAVE IN', 'TOTAL IN', 'REJECTION OUT TO RPS', 'QC REWORK OUT', 'FG OUT', 'TOTAL OUT', 'CURRENT STOCK'] :
      ['INHOUSE OPENING STOCK', 'VENDOR OPENING STOCK', 'TRIMMING VENDOR IN', 'QC REWORK IN', 'MOLD IN', 'METAL STORE IN', 'EXTRUSION IN', 'TOTAL IN', 'QC OUT', 'VENDOR OUT', 'REJECTION OUT TO RPS', 'TOTAL OUT', 'IN HOUSE STOCK', 'VENDOR STOCK', 'TOTAL STOCK'];

    const groupSize = groupHeaders.length;
    const firstColHeader = activeTab === 'mini-store' ? 'ITEM ID' : (activeTab === 'quality' || activeTab === 'fg-store' ? 'Part No. & Name' : 'Item ID');

    // 3. Group transactions by part name for efficient lookup
    const transByPart = new Map<string, Transaction[]>();
    transactions.forEach(t => {
      const key = t.partName.toString().trim().toLowerCase();
      if (!transByPart.has(key)) transByPart.set(key, []);
      transByPart.get(key)!.push(t);
    });

    const uniqueTransPartNames = Array.from(transByPart.keys());
    const partialMatchCache = new Map<string, string[]>();

    const getFilteredTrans = (partName: string) => {
      const sPartKey = partName.toString().trim().toLowerCase();
      let filtered = transByPart.get(sPartKey) || [];
      if (filtered.length === 0) {
        if (!partialMatchCache.has(sPartKey)) {
          const matchedKeys = uniqueTransPartNames.filter(tKey => 
            tKey.includes(sPartKey) || sPartKey.includes(tKey)
          );
          partialMatchCache.set(sPartKey, matchedKeys);
        }
        const matchedKeys = partialMatchCache.get(sPartKey)!;
        filtered = matchedKeys.flatMap(k => transByPart.get(k) || []);
      }
      return filtered;
    };

    const calcTotals = (transList: Transaction[]) => {
      let moldIn = 0, reworkIn = 0, qcOut = 0, rejectionOut = 0;
      let fgReworkIn = 0, metalStoreIn = 0, customerRejectionIn = 0, oilSealTrimmingIn = 0, trimmingIn = 0, extrusionIn = 0;
      let rejectionOutToRps = 0, metalStoreOut = 0, oilSealTrimmingOut = 0, trimmingOut = 0, fgOut = 0, extrusionOut = 0;
      let trimmingVendorIn = 0, trimmingQcReworkIn = 0, trimmingMoldingIn = 0, trimmingMetalStoreIn = 0, trimmingExtrusionIn = 0;
      let trimmingQcOut = 0, trimmingVendorOut = 0, trimmingRejectionOutToRps = 0;
      let qcIn = 0, autoClaveIn = 0, qcReworkOut = 0;
      let compoundIn = 0, moldReturnIn = 0, vendorOut = 0, injectOut = 0, oilSealOut = 0, moldOut = 0, autoClaveOut = 0, labOut = 0;
      let chemicalStoreIn = 0, phosphateIn = 0, injcMoldOut = 0, hvcmOut = 0, rejectionOutToMetalStore = 0;
      let autoClaveProdIn = 0, autoClaveMiniStoreIn = 0, autoClaveMetalIn = 0, autoClaveReworkIn = 0, autoClaveRejectionOut = 0, autoClaveMetalOut = 0;
      let phosphateOutToBonding = 0;
      let totalIn = 0, totalOut = 0;

      transList.forEach(t => {
        const type = t.type.toLowerCase();
        const qty = t.quantity;
        if (activeTab === 'bonding') {
          if (type.includes('metal') && type.includes('in')) metalStoreIn += qty;
          else if (type.includes('chemical') && type.includes('in')) chemicalStoreIn += qty;
          else if (type.includes('phosphate') && type.includes('in')) phosphateIn += qty;
          else if (type.includes('mold') && type.includes('in')) moldIn += qty;
          else if (type.includes('injc') && type.includes('out')) injcMoldOut += qty;
          else if (type.includes('oil') && type.includes('seal') && type.includes('out')) oilSealOut += qty;
          else if (type.includes('hvcm') && type.includes('out')) hvcmOut += qty;
          else if (type.includes('rejection') && type.includes('out')) rejectionOutToMetalStore += qty;
          else if (type.includes('rejection') && !type.includes('in')) rejectionOutToMetalStore += qty;
        } else if (activeTab === 'auto-clave') {
          if (type.includes('prod') && type.includes('in')) autoClaveProdIn += qty;
          else if (type.includes('mini') && type.includes('store') && type.includes('in')) autoClaveMiniStoreIn += qty;
          else if (type.includes('metal') && type.includes('in')) autoClaveMetalIn += qty;
          else if (type.includes('rework') && type.includes('in')) autoClaveReworkIn += qty;
          else if (type.includes('metal') && type.includes('out')) autoClaveMetalOut += qty;
          else if (type.includes('reject') && type.includes('rps')) autoClaveRejectionOut += qty;
        } else if (activeTab === 'phosphate') {
          if (type.includes('metal') && type.includes('in')) metalStoreIn += qty;
          else if (type.includes('chemical') && type.includes('in')) chemicalStoreIn += qty;
          else if (type.includes('bonding') && type.includes('out')) phosphateOutToBonding += qty;
          else if (type.includes('reject') && type.includes('rps')) rejectionOutToRps += qty;
        } else if (activeTab === 'oil-seal' || activeTab === 'extrusion') {
          if (type.includes('mold') && type.includes('in')) moldIn += qty;
          else if (type.includes('re work in') || type.includes('rework in')) reworkIn += qty;
          else if (type.includes('qc out')) qcOut += qty;
          else if (type.includes('rejection to rps out') || type.includes('rejection')) rejectionOut += qty;
        } else if (activeTab === 'molding') {
          if (type.includes('rejection out to rps') || (type.includes('rejection') && type.includes('rps'))) rejectionOutToRps += qty;
          else if (type.includes('oil seal trimming out')) oilSealTrimmingOut += qty;
          else if (type.includes('trimming out') && !type.includes('oil seal')) trimmingOut += qty;
        } else if (activeTab === 'mini-store') {
          const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
          const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

          if (type.includes('compound') && isRecd) compoundIn += qty;
          else if (type.includes('mold return') && isRecd) moldReturnIn += qty;
          else if (type.includes('vendor') && isSent) vendorOut += qty;
          else if (type.includes('inject') && isSent) injectOut += qty;
          else if (type.includes('oil seal') && isSent) oilSealOut += qty;
          else if (type.includes('mold') && isSent) moldOut += qty;
          else if (type.includes('extru') && isSent) extrusionOut += qty;
          else if (type.includes('auto') && type.includes('clave') && isSent) autoClaveOut += qty;
          else if (type.includes('lab') && isSent) labOut += qty;
          else if (type.includes('reject') && (isSent || !isRecd)) rejectionOutToRps += qty;

          if (isRecd) totalIn += qty;
          else if (isSent || type.includes('reject')) totalOut += qty;
        } else if (activeTab === 'quality') {
          const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
          const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

          if (type.includes('fg rework') && isRecd) fgReworkIn += qty;
          else if (type.includes('metal') && isRecd) metalStoreIn += qty;
          else if (type.includes('customer') && isRecd) customerRejectionIn += qty;
          else if (type.includes('oil') && type.includes('seal') && type.includes('trim') && isRecd) oilSealTrimmingIn += qty;
          else if (type.includes('trim') && isRecd) trimmingIn += qty;
          else if (type.includes('extru') && isRecd) extrusionIn += qty;
          else if (type.includes('reject') && (isSent || !isRecd)) rejectionOutToRps += qty;
          else if (type.includes('metal') && isSent) metalStoreOut += qty;
          else if (type.includes('oil') && type.includes('seal') && type.includes('trim') && isSent) oilSealTrimmingOut += qty;
          else if (type.includes('trim') && isSent) trimmingOut += qty;
          else if (type.includes('fg') && isSent) fgOut += qty;
          else if (type.includes('extru') && isSent) extrusionOut += qty;

          if (isRecd) totalIn += qty;
          else if (isSent || type.includes('reject')) totalOut += qty;
        } else if (activeTab === 'fg-store') {
          const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
          const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

          if (type.includes('customer') && isRecd) customerRejectionIn += qty;
          else if (type.includes('qc') && isRecd) qcIn += qty;
          else if (type.includes('rework') && isRecd) reworkIn += qty;
          else if (type.includes('auto') && type.includes('clave') && isRecd) autoClaveIn += qty;
          else if (type.includes('reject') && type.includes('rps') && isSent) rejectionOutToRps += qty;
          else if (type.includes('qc') && type.includes('rework') && isSent) qcReworkOut += qty;
          else if (type.includes('fg') && isSent) fgOut += qty;
        } else if (activeTab === 'trimming') {
          const isRecd = type.includes('in') || type.includes('recd') || type.includes('received') || type.includes('ok');
          const isSent = type.includes('out') || type.includes('sent') || type.includes('issue');

          if (type.includes('vendor') && isRecd) trimmingVendorIn += qty;
          else if (type.includes('qc rework')) trimmingQcReworkIn += qty;
          else if (type.includes('mold') && isRecd) trimmingMoldingIn += qty;
          else if (type.includes('metal') && isRecd) trimmingMetalStoreIn += qty;
          else if (type.includes('extru') && isRecd) trimmingExtrusionIn += qty;
          else if (type.includes('qc') && isSent) trimmingQcOut += qty;
          else if (type.includes('vendor') && isSent) trimmingVendorOut += qty;
          else if (type.includes('reject') && (isSent || !isRecd)) trimmingRejectionOutToRps += qty;

          if (isRecd) totalIn += qty;
          else if (isSent || type.includes('reject')) totalOut += qty;
        }
      });

      if (activeTab === 'bonding') {
        totalIn = metalStoreIn + chemicalStoreIn + phosphateIn + moldIn;
        totalOut = injcMoldOut + oilSealOut + hvcmOut + rejectionOutToMetalStore;
      } else if (activeTab === 'auto-clave') {
        totalIn = autoClaveProdIn + autoClaveMiniStoreIn + autoClaveMetalIn + autoClaveReworkIn;
        totalOut = autoClaveRejectionOut + autoClaveMetalOut;
      } else if (activeTab === 'phosphate') {
        totalIn = metalStoreIn + chemicalStoreIn;
        totalOut = phosphateOutToBonding + rejectionOutToRps;
      } else if (activeTab === 'oil-seal' || activeTab === 'extrusion') {
        totalIn = moldIn + reworkIn;
        totalOut = qcOut + rejectionOut;
      } else if (activeTab === 'molding') {
        totalOut = rejectionOutToRps + oilSealTrimmingOut + trimmingOut;
      } else if (activeTab === 'quality') {
        totalIn = fgReworkIn + metalStoreIn + customerRejectionIn + oilSealTrimmingIn + trimmingIn + extrusionIn;
        totalOut = rejectionOutToRps + metalStoreOut + oilSealTrimmingOut + trimmingOut + fgOut + extrusionOut;
      } else if (activeTab === 'mini-store') {
        totalIn = compoundIn + moldReturnIn;
        totalOut = vendorOut + injectOut + oilSealOut + moldOut + extrusionOut + autoClaveOut + labOut + rejectionOutToRps;
      } else if (activeTab === 'fg-store') {
        totalIn = customerRejectionIn + qcIn + reworkIn + autoClaveIn;
        totalOut = rejectionOutToRps + qcReworkOut + fgOut;
      } else if (activeTab === 'trimming') {
        totalIn = trimmingVendorIn + trimmingQcReworkIn + trimmingMoldingIn + trimmingMetalStoreIn + trimmingExtrusionIn;
        totalOut = trimmingQcOut + trimmingVendorOut + trimmingRejectionOutToRps;
      }

      return {
        moldIn, reworkIn, qcOut, rejectionOut,
        fgReworkIn, metalStoreIn, customerRejectionIn, oilSealTrimmingIn, trimmingIn, extrusionIn,
        rejectionOutToRps, metalStoreOut, oilSealTrimmingOut, trimmingOut, fgOut, extrusionOut,
        trimmingVendorIn, trimmingQcReworkIn, trimmingMoldingIn, trimmingMetalStoreIn, trimmingExtrusionIn,
        trimmingQcOut, trimmingVendorOut, trimmingRejectionOutToRps,
        qcIn, autoClaveIn, qcReworkOut,
        compoundIn, moldReturnIn, vendorOut, injectOut, oilSealOut, moldOut, autoClaveOut, labOut,
        chemicalStoreIn, phosphateIn, injcMoldOut, hvcmOut, rejectionOutToMetalStore,
        autoClaveProdIn, autoClaveMiniStoreIn, autoClaveMetalIn, autoClaveReworkIn, autoClaveRejectionOut, autoClaveMetalOut,
        phosphateOutToBonding,
        totalIn, totalOut
      };
    };

    // 4. Build the AOA (Array of Arrays)
    const sheetData: any[] = [];
    
    // Row 1: Main Title
    const mainTitle = `${tabTitle} Stock Report - ${selectedMonth}`;
    sheetData.push([mainTitle]);

    // Row 2: Group Headers (TOTAL, Day 1, Day 2, etc.)
    const row2 = ['', '']; // Empty for Item ID and Job # col
    row2.push('TOTAL');
    for (let i = 1; i < groupSize; i++) row2.push(''); // Fill for merge
    for (const day of sortedActiveDays) {
      row2.push(format(day, 'EEEE, MMMM d, yyyy'));
      for (let i = 1; i < groupSize; i++) row2.push(''); // Fill for merge
    }
    sheetData.push(row2);

    // Filter openingStocks to only include items with activity
    const activeOpeningStocks = openingStocks.filter(stock => {
      const opening = stock.monthlyStocks[selectedMonth] || 0;
      const filteredTrans = getFilteredTrans(stock.partName);
      const monthTrans = filteredTrans.filter(t => t.parsedDate && t.parsedDate >= startOfM && t.parsedDate <= endOfM);
      return opening !== 0 || monthTrans.length > 0;
    });

    // Row 3: Totals (Formulas)
    const row3: any[] = ['Total', ''];
    const totalCols = 2 + groupSize + (sortedActiveDays.length * groupSize);
    const lastDataRow = activeOpeningStocks.length + 4;
    for (let i = 2; i < totalCols; i++) {
      const colLetter = XLSX.utils.encode_col(i);
      row3.push({ f: `SUM(${colLetter}5:${colLetter}${lastDataRow})` });
    }
    sheetData.push(row3);

    // Row 4: Column Headers
    const row4 = [firstColHeader, 'Job #'];
    row4.push(...groupHeaders); // For TOTAL group
    for (let i = 0; i < sortedActiveDays.length; i++) {
      row4.push(...groupHeaders); // For each Day group
    }
    sheetData.push(row4);

    // Row 5+: Data Rows
    activeOpeningStocks.forEach((stock, index) => {
      const summaryItem = allItemsWithActivity.find(item => item.itemId === stock.partName);
      const jobId = summaryItem?.jobId || '';
      const row: any[] = [stock.partName, jobId];
      const rowNum = index + 5; // Excel rows are 1-indexed, data starts at Row 5
      const opening = stock.monthlyStocks[selectedMonth] || 0;
      const filteredTrans = getFilteredTrans(stock.partName);
      const sPartKey = stock.partName.toString().trim().toLowerCase();
      const vendorOpening = activeTab === 'trimming' ? (vendorStocksByPart.get(sPartKey) || 0) : 0;

      // --- TOTAL Section ---
      // For each transaction type in the TOTAL section, we want a formula that sums the same type across all days
      const getSumFormula = (offset: number) => {
        const dailyCols = sortedActiveDays.map((_, dIdx) => {
          const groupStartCol = 2 + groupSize + (dIdx * groupSize);
          return `${XLSX.utils.encode_col(groupStartCol + offset)}${rowNum}`;
        });
        return { f: dailyCols.join('+') };
      };

      if (activeTab === 'bonding') {
        row.push(
          opening, 
          getSumFormula(1), // Metal Store In
          getSumFormula(2), // Chemical Store In
          getSumFormula(3), // Phosphate In
          getSumFormula(4), // Mold In
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}+${XLSX.utils.encode_col(5)}${rowNum}+${XLSX.utils.encode_col(6)}${rowNum}` }, // Total IN (C+D+E+F+G)
          getSumFormula(6), // Injc Mold Out
          getSumFormula(7), // Oil Seal Out
          getSumFormula(8), // Hvcm Out
          getSumFormula(9), // Rejection Out
          { f: `${XLSX.utils.encode_col(8)}${rowNum}+${XLSX.utils.encode_col(9)}${rowNum}+${XLSX.utils.encode_col(10)}${rowNum}+${XLSX.utils.encode_col(11)}${rowNum}` }, // Total OUT (I+J+K+L)
          { f: `${XLSX.utils.encode_col(7)}${rowNum}-${XLSX.utils.encode_col(12)}${rowNum}` }  // Current Stock (H-M)
        );
      } else if (activeTab === 'auto-clave') {
        row.push(
          opening, 
          getSumFormula(1), // Prod In
          getSumFormula(2), // Mini Store In
          getSumFormula(3), // Metal In
          getSumFormula(4), // Rework In
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}+${XLSX.utils.encode_col(5)}${rowNum}+${XLSX.utils.encode_col(6)}${rowNum}` }, // Total IN
          getSumFormula(6), // Rejection Out
          getSumFormula(7), // Metal Out
          { f: `${XLSX.utils.encode_col(8)}${rowNum}+${XLSX.utils.encode_col(9)}${rowNum}` }, // Total OUT
          { f: `${XLSX.utils.encode_col(7)}${rowNum}-${XLSX.utils.encode_col(10)}${rowNum}` }  // Current Stock
        );
      } else if (activeTab === 'phosphate') {
        row.push(
          opening,
          getSumFormula(1), // Metal Store In
          getSumFormula(2), // Chemical Store In
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}` }, // Total IN
          getSumFormula(4), // Phosphate Out To Bonding
          getSumFormula(5), // Rejection Out To RPS
          { f: `${XLSX.utils.encode_col(6)}${rowNum}+${XLSX.utils.encode_col(7)}${rowNum}` }, // Total OUT
          { f: `${XLSX.utils.encode_col(5)}${rowNum}-${XLSX.utils.encode_col(8)}${rowNum}` }  // Current Stock
        );
      } else if (activeTab === 'oil-seal' || activeTab === 'extrusion') {
        row.push(
          opening, 
          getSumFormula(1), // Mold IN
          getSumFormula(2), // Re Work IN
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}` }, // Total IN
          getSumFormula(4), // QC Out
          getSumFormula(5), // Rejection on OUT
          { f: `${XLSX.utils.encode_col(6)}${rowNum}+${XLSX.utils.encode_col(7)}${rowNum}` }, // Total OUT
          { f: `${XLSX.utils.encode_col(5)}${rowNum}-${XLSX.utils.encode_col(8)}${rowNum}` }  // Current Stock
        );
      } else if (activeTab === 'molding') {
        row.push(
          getSumFormula(0), // Rejection OUT to RPS
          getSumFormula(1), // Oil Seal Trimming OUT
          getSumFormula(2), // Trimming OUT
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}` } // Total OUT
        );
      } else if (activeTab === 'quality') {
        row.push(
          opening, 
          getSumFormula(1), getSumFormula(2), getSumFormula(3), getSumFormula(4), getSumFormula(5), getSumFormula(6), // INs
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}+${XLSX.utils.encode_col(5)}${rowNum}+${XLSX.utils.encode_col(6)}${rowNum}+${XLSX.utils.encode_col(7)}${rowNum}+${XLSX.utils.encode_col(8)}${rowNum}` },
          getSumFormula(8), getSumFormula(9), getSumFormula(10), getSumFormula(11), getSumFormula(12), getSumFormula(13), // OUTs
          { f: `${XLSX.utils.encode_col(10)}${rowNum}+${XLSX.utils.encode_col(11)}${rowNum}+${XLSX.utils.encode_col(12)}${rowNum}+${XLSX.utils.encode_col(13)}${rowNum}+${XLSX.utils.encode_col(14)}${rowNum}+${XLSX.utils.encode_col(15)}${rowNum}` },
          { f: `${XLSX.utils.encode_col(9)}${rowNum}-${XLSX.utils.encode_col(16)}${rowNum}` }
        );
      } else if (activeTab === 'mini-store') {
        row.push(
          opening,
          getSumFormula(1), getSumFormula(2), // COMPOUND IN, MOLD RETURN IN
          { f: `${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}` }, // Total IN
          getSumFormula(4), getSumFormula(5), getSumFormula(6), getSumFormula(7), getSumFormula(8), getSumFormula(9), getSumFormula(10), getSumFormula(11), // OUTs
          { f: `${XLSX.utils.encode_col(6)}${rowNum}+${XLSX.utils.encode_col(7)}${rowNum}+${XLSX.utils.encode_col(8)}${rowNum}+${XLSX.utils.encode_col(9)}${rowNum}+${XLSX.utils.encode_col(10)}${rowNum}+${XLSX.utils.encode_col(11)}${rowNum}+${XLSX.utils.encode_col(12)}${rowNum}+${XLSX.utils.encode_col(13)}${rowNum}` }, // Total OUT
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(5)}${rowNum}-${XLSX.utils.encode_col(14)}${rowNum}` } // Current Stock
        );
      } else if (activeTab === 'fg-store') {
        row.push(
          opening,
          getSumFormula(1), getSumFormula(2), getSumFormula(3), getSumFormula(4), // INs
          { f: `${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}+${XLSX.utils.encode_col(5)}${rowNum}+${XLSX.utils.encode_col(6)}${rowNum}` }, // Total IN (D+E+F+G)
          getSumFormula(6), getSumFormula(7), getSumFormula(8), // OUTs
          { f: `${XLSX.utils.encode_col(8)}${rowNum}+${XLSX.utils.encode_col(9)}${rowNum}+${XLSX.utils.encode_col(10)}${rowNum}` }, // Total OUT (I+J+K)
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(7)}${rowNum}-${XLSX.utils.encode_col(11)}${rowNum}` } // Current Stock (C+H-L)
        );
      } else {
        // Trimming
        row.push(
          opening, 
          vendorOpening,
          getSumFormula(2), getSumFormula(3), getSumFormula(4), getSumFormula(5), getSumFormula(6), // INs
          { f: `${XLSX.utils.encode_col(2)}${rowNum}+${XLSX.utils.encode_col(4)}${rowNum}+${XLSX.utils.encode_col(5)}${rowNum}+${XLSX.utils.encode_col(6)}${rowNum}+${XLSX.utils.encode_col(7)}${rowNum}+${XLSX.utils.encode_col(8)}${rowNum}` }, // Total IN
          getSumFormula(8), getSumFormula(9), getSumFormula(10), // OUTs
          { f: `${XLSX.utils.encode_col(10)}${rowNum}+${XLSX.utils.encode_col(11)}${rowNum}+${XLSX.utils.encode_col(12)}${rowNum}` }, // Total OUT
          { f: `${XLSX.utils.encode_col(9)}${rowNum}-${XLSX.utils.encode_col(13)}${rowNum}` }, // IN HOUSE STOCK
          { f: `${XLSX.utils.encode_col(3)}${rowNum}+${XLSX.utils.encode_col(11)}${rowNum}-${XLSX.utils.encode_col(4)}${rowNum}` }, // VENDOR STOCK
          { f: `${XLSX.utils.encode_col(14)}${rowNum}+${XLSX.utils.encode_col(15)}${rowNum}` } // TOTAL STOCK
        );
      }

      // --- Daily Sections ---
      sortedActiveDays.forEach((day, dIdx) => {
        const startOfD = startOfDay(day);
        const endOfD = endOfDay(day);
        const dayTrans = filteredTrans.filter(t => t.parsedDate && t.parsedDate >= startOfD && t.parsedDate <= endOfD);
        const dTotals = calcTotals(dayTrans);
        
        const groupStartCol = 2 + groupSize + (dIdx * groupSize);
        const prevGroupEndCol = groupStartCol - 1;
        
        // Opening for Day N is Current Stock of Day N-1 (or TOTAL Opening for Day 1)
        const dOpening = dIdx === 0 ? opening : { f: `${XLSX.utils.encode_col(prevGroupEndCol)}${rowNum}` };

        if (activeTab === 'bonding') {
          row.push(
            dOpening, 
            dTotals.metalStoreIn, 
            dTotals.chemicalStoreIn, 
            dTotals.phosphateIn, 
            dTotals.moldIn, 
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+3)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+4)}${rowNum}` },
            dTotals.injcMoldOut, 
            dTotals.oilSealOut, 
            dTotals.hvcmOut, 
            dTotals.rejectionOutToMetalStore, 
            { f: `${XLSX.utils.encode_col(groupStartCol+6)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+7)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+8)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+9)}${rowNum}` },
            { f: `${XLSX.utils.encode_col(groupStartCol+5)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+10)}${rowNum}` }
          );
        } else if (activeTab === 'auto-clave') {
          row.push(
            dOpening, 
            dTotals.autoClaveProdIn, 
            dTotals.autoClaveMiniStoreIn, 
            dTotals.autoClaveMetalIn, 
            dTotals.autoClaveReworkIn, 
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+3)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+4)}${rowNum}` },
            dTotals.autoClaveRejectionOut, 
            dTotals.autoClaveMetalOut, 
            { f: `${XLSX.utils.encode_col(groupStartCol+6)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+7)}${rowNum}` },
            { f: `${XLSX.utils.encode_col(groupStartCol+5)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+8)}${rowNum}` }
          );
        } else if (activeTab === 'phosphate') {
          row.push(
            dOpening, 
            dTotals.metalStoreIn, 
            dTotals.chemicalStoreIn, 
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}` },
            dTotals.phosphateOutToBonding, 
            dTotals.rejectionOutToRps, 
            { f: `${XLSX.utils.encode_col(groupStartCol+4)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+5)}${rowNum}` },
            { f: `${XLSX.utils.encode_col(groupStartCol+3)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+6)}${rowNum}` }
          );
        } else if (activeTab === 'oil-seal' || activeTab === 'extrusion') {
          row.push(
            dOpening, 
            dTotals.moldIn, 
            dTotals.reworkIn, 
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}` },
            dTotals.qcOut, 
            dTotals.rejectionOut, 
            { f: `${XLSX.utils.encode_col(groupStartCol+4)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+5)}${rowNum}` },
            { f: `${XLSX.utils.encode_col(groupStartCol+3)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+6)}${rowNum}` }
          );
        } else if (activeTab === 'molding') {
          row.push(
            dTotals.rejectionOutToRps,
            dTotals.oilSealTrimmingOut,
            dTotals.trimmingOut,
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}` }
          );
        } else if (activeTab === 'quality') {
          row.push(
            dOpening, dTotals.fgReworkIn, dTotals.metalStoreIn, dTotals.customerRejectionIn, dTotals.oilSealTrimmingIn, dTotals.trimmingIn, dTotals.extrusionIn,
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+3)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+4)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+5)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+6)}${rowNum}` },
            dTotals.rejectionOutToRps, dTotals.metalStoreOut, dTotals.oilSealTrimmingOut, dTotals.trimmingOut, dTotals.fgOut, dTotals.extrusionOut,
            { f: `${XLSX.utils.encode_col(groupStartCol+8)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+9)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+10)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+11)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+12)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+13)}${rowNum}` },
            { f: `${XLSX.utils.encode_col(groupStartCol+7)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+14)}${rowNum}` }
          );
        } else if (activeTab === 'mini-store') {
          row.push(
            dOpening, dTotals.compoundIn, dTotals.moldReturnIn,
            { f: `${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}` }, // Total IN
            dTotals.vendorOut, dTotals.injectOut, dTotals.oilSealOut, dTotals.moldOut, dTotals.extrusionOut, dTotals.autoClaveOut, dTotals.labOut, dTotals.rejectionOutToRps,
            { f: `${XLSX.utils.encode_col(groupStartCol+4)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+5)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+6)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+7)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+8)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+9)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+10)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+11)}${rowNum}` }, // Total OUT
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+3)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+12)}${rowNum}` } // Total Stock
          );
        } else if (activeTab === 'fg-store') {
          row.push(
            dOpening, dTotals.customerRejectionIn, dTotals.qcIn, dTotals.reworkIn, dTotals.autoClaveIn,
            { f: `${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+3)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+4)}${rowNum}` }, // Total IN
            dTotals.rejectionOutToRps, dTotals.qcReworkOut, dTotals.fgOut,
            { f: `${XLSX.utils.encode_col(groupStartCol+6)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+7)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+8)}${rowNum}` }, // Total OUT
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+5)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+9)}${rowNum}` } // Current Stock
          );
        } else {
          // Trimming
          const dInhouseOpening = dIdx === 0 ? opening : { f: `${XLSX.utils.encode_col(prevGroupEndCol - 2)}${rowNum}` };
          const dVendorOpening = dIdx === 0 ? vendorOpening : { f: `${XLSX.utils.encode_col(prevGroupEndCol - 1)}${rowNum}` };

          row.push(
            dInhouseOpening, 
            dVendorOpening,
            dTotals.trimmingVendorIn, dTotals.trimmingQcReworkIn, dTotals.trimmingMoldingIn, dTotals.trimmingMetalStoreIn, dTotals.trimmingExtrusionIn,
            { f: `${XLSX.utils.encode_col(groupStartCol)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+3)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+4)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+5)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+6)}${rowNum}` }, // Total IN
            dTotals.trimmingQcOut, dTotals.trimmingVendorOut, dTotals.trimmingRejectionOutToRps,
            { f: `${XLSX.utils.encode_col(groupStartCol+8)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+9)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+10)}${rowNum}` }, // Total OUT
            { f: `${XLSX.utils.encode_col(groupStartCol+7)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+11)}${rowNum}` }, // IN HOUSE STOCK
            { f: `${XLSX.utils.encode_col(groupStartCol+1)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+9)}${rowNum}-${XLSX.utils.encode_col(groupStartCol+2)}${rowNum}` }, // VENDOR STOCK
            { f: `${XLSX.utils.encode_col(groupStartCol+12)}${rowNum}+${XLSX.utils.encode_col(groupStartCol+13)}${rowNum}` } // TOTAL STOCK
          );
        }
      });

      sheetData.push(row);
    });

    const worksheet = XLSX.utils.aoa_to_sheet(sheetData);

    // 5. Merges
    const merges: XLSX.Range[] = [];
    // Merge Row 1 Title
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: totalCols - 1 } });
    // Merge Row 2 Groups
    merges.push({ s: { r: 1, c: 2 }, e: { r: 1, c: 2 + groupSize - 1 } }); // TOTAL group
    for (let i = 0; i < sortedActiveDays.length; i++) {
      const startCol = 2 + (i + 1) * groupSize;
      merges.push({ s: { r: 1, c: startCol }, e: { r: 1, c: startCol + groupSize - 1 } });
    }
    worksheet['!merges'] = merges;

    // 6. Styling
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let R = range.s.r; R <= range.e.r; ++R) {
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!worksheet[addr]) continue;
        
        // Determine if this column is the end of a group (for bold right border)
        const isGroupEnd = C === 1 || (C > 1 && (C - 1) % groupSize === 0);

        worksheet[addr].s = {
          border: { 
            top: { style: 'thin' }, 
            bottom: { style: 'thin' }, 
            left: { style: 'thin' }, 
            right: { style: isGroupEnd ? 'medium' : 'thin' } 
          },
          alignment: { 
            horizontal: C === 0 ? 'left' : 'center', 
            vertical: 'center',
            wrapText: R === 3 // Wrap text for column headers
          }
        };

        if (R === 0) { // Main Title
          worksheet[addr].s.font = { bold: true, size: 14 };
          worksheet[addr].s.fill = { fgColor: { rgb: 'E9ECEF' } };
        } else if (R === 1) { // Group Headers (TOTAL, Dates)
          worksheet[addr].s.font = { bold: true, size: 12 };
          worksheet[addr].s.fill = { fgColor: { rgb: C <= groupSize ? 'F8F9FA' : 'D1E7DD' } };
        } else if (R === 2 || R === 3) { // Totals & Column Headers
          worksheet[addr].s.font = { bold: true };
          worksheet[addr].s.fill = { fgColor: { rgb: R === 2 ? 'F1F3F5' : 'FFFFFF' } };
          
          // Color coding for headers based on text
          const headerCell = worksheet[XLSX.utils.encode_cell({ r: 3, c: C })];
          const hText = (headerCell?.v || '').toString().toUpperCase();
          
          if (hText.includes('OPENING')) {
            worksheet[addr].s.font.color = { rgb: 'D35400' }; // Orange/Brown
          } else if (hText === 'TOTAL IN') {
            worksheet[addr].s.font.color = { rgb: '27AE60' }; // Green
          } else if (hText === 'TOTAL OUT') {
            worksheet[addr].s.font.color = { rgb: 'C0392B' }; // Red
          } else if (hText.includes('STOCK')) {
            worksheet[addr].s.font.color = { rgb: '2980B9' }; // Blue
          }
        }
      }
    }

    worksheet['!cols'] = [{ wch: 45 }, ...new Array(totalCols - 1).fill({ wch: 10 })]; // Reduced width to 10
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Daily Breakdown');

    // Add Daily Summary Sheet
    const initialOpening = activeOpeningStocks.reduce((sum, item) => sum + (item.monthlyStocks[selectedMonth] || 0), 0);
    const initialVendorOpening = activeTab === 'trimming' ? activeOpeningStocks.reduce((sum, item) => {
      const sPartKey = item.partName.toString().trim().toLowerCase();
      return sum + (vendorStocksByPart.get(sPartKey) || 0);
    }, 0) : 0;
    addDailySummarySheetToWorkbook(workbook, tabTitle, selectedMonth, activeTab, sortedDailySummary, dailyTotals, initialOpening, initialVendorOpening);

    const fileName = `${tabTitle.replace(/\s+/g, '_')}_Daily_Report_${selectedMonth.replace(/\s+/g, '_')}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  const handleExportDailySummaryTable = () => {
    if (sortedDailySummary.length === 0) return;

    const workbook = XLSX.utils.book_new();
    const tabTitle = activeTab === 'molding' ? 'Molding' : 
                     activeTab === 'oil-seal' || activeTab === 'bonding' || activeTab === 'phosphate' || activeTab === 'auto-clave' || activeTab === 'extrusion' ? 
                     (activeTab === 'oil-seal' ? 'Oil Seal Trimming' : activeTab === 'bonding' ? 'Bonding' : activeTab === 'phosphate' ? 'Phosphate' : activeTab === 'auto-clave' ? 'Auto Clave' : 'Extrusion') : 
                     activeTab === 'quality' ? 'Quality' : 
                     activeTab === 'mini-store' ? 'Mini Store' :
                     activeTab === 'fg-store' ? 'FG Store' :
                     'Trimming';
    
    const initialOpening = summaryData.reduce((sum, item) => sum + item.openingStock, 0);
    const initialVendorOpening = activeTab === 'trimming' || activeTab === 'extrusion' ? summaryData.reduce((sum, item) => sum + (item.vendorOpeningStock || 0), 0) : 0;
    addDailySummarySheetToWorkbook(workbook, tabTitle, selectedMonth, activeTab, sortedDailySummary, dailyTotals, initialOpening, initialVendorOpening);

    const fileName = `${tabTitle.replace(/\s+/g, '_')}_Daily_Summary_${selectedMonth.replace(/\s+/g, '_')}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  if (loading && openingStocks.length === 0) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-4">
        <RefreshCw className="w-12 h-12 text-blue-600 animate-spin mb-4" />
        <p className="text-slate-600 font-medium animate-pulse">Loading Inventory Data...</p>
      </div>
    );
  }

  return (
    <div className="flex h-screen bg-slate-50 text-slate-900 font-sans overflow-hidden">
      {/* Persistent Sidebar */}
      <motion.aside
        initial={false}
        animate={{ width: isSidebarCollapsed ? 80 : 280 }}
        className="flex flex-col bg-white border-r border-slate-200 z-[150] relative"
      >
        {/* Sidebar Header */}
        <div className="p-6 flex items-center justify-between overflow-hidden">
          <div className="flex items-center gap-3 overflow-hidden">
            <div className="bg-gradient-to-br from-indigo-500 to-purple-600 p-2 rounded-xl shadow-lg shadow-indigo-100 shrink-0">
              <Activity className="w-6 h-6 text-white" />
            </div>
            {!isSidebarCollapsed && (
              <motion.div
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                className="overflow-hidden whitespace-nowrap"
              >
                <h2 className="text-lg font-black text-slate-800 leading-none tracking-tight">ProDash</h2>
                <span className="text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">Analytics</span>
              </motion.div>
            )}
          </div>
          <button
            onClick={() => setIsSidebarCollapsed(!isSidebarCollapsed)}
            className="p-1.5 hover:bg-slate-50 rounded-lg transition-colors text-slate-300 hover:text-slate-600 border border-slate-100"
          >
            {isSidebarCollapsed ? <ChevronRight className="w-4 h-4" /> : <ChevronLeft className="w-4 h-4" />}
          </button>
        </div>

        {/* Sidebar Navigation */}
        <div className="flex-1 overflow-y-auto px-4 py-6 scrollbar-hide">
          <div className="space-y-1">
            <div className={cn(
              "px-3 mb-3",
              isSidebarCollapsed ? "hidden" : "block"
            )}>
              <span className="text-[10px] font-black text-indigo-400 uppercase tracking-widest leading-none">Main Navigation</span>
            </div>
            {MENU_ITEMS.map((item) => (
              <button
                key={item.id}
                onClick={() => setActiveTab(item.id as any)}
                className={cn(
                  "w-full flex items-center gap-3 px-3 py-3 rounded-xl transition-all group overflow-hidden relative",
                  activeTab === item.id 
                    ? "bg-indigo-50/80 text-indigo-600 font-bold" 
                    : "hover:bg-slate-50 text-slate-500 hover:text-slate-900"
                )}
                title={isSidebarCollapsed ? item.name : ""}
              >
                <div className={cn(
                  "p-1.5 rounded-lg transition-colors shrink-0",
                  activeTab === item.id ? "text-indigo-600" : "text-slate-400 group-hover:text-slate-600"
                )}>
                  {item.icon}
                </div>
                {!isSidebarCollapsed && (
                  <motion.span
                    initial={{ opacity: 0, x: -5 }}
                    animate={{ opacity: 1, x: 0 }}
                    className="text-[13px] tracking-tight whitespace-nowrap uppercase font-bold"
                  >
                    {item.name}
                  </motion.span>
                )}
                {activeTab === item.id && !isSidebarCollapsed && (
                  <div className="ml-auto">
                    <ChevronDown className="w-3 h-3 opacity-40" />
                  </div>
                )}
              </button>
            ))}
          </div>

          <div className="mt-8 space-y-1">
            <div className={cn(
              "px-3 mb-3",
              isSidebarCollapsed ? "hidden" : "block"
            )}>
              <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest leading-none">System</span>
            </div>
            <button className="w-full flex items-center gap-3 px-3 py-3 rounded-xl hover:bg-slate-50 text-slate-500 hover:text-slate-900 transition-all overflow-hidden group">
              <div className="p-1.5 text-slate-400 group-hover:text-slate-600 shrink-0">
                <Settings className="w-4 h-4" />
              </div>
              {!isSidebarCollapsed && <span className="text-[13px] tracking-tight transition-margin">Settings</span>}
            </button>
          </div>
        </div>

        {/* Sidebar Footer */}
        <div className="p-4 bg-slate-50/50">
          <div className={cn(
            "flex items-center gap-3 px-3 py-3 rounded-2xl bg-white border border-slate-100 shadow-sm transition-all",
            isSidebarCollapsed ? "justify-center px-1" : ""
          )}>
            <div className="w-8 h-8 rounded-lg bg-indigo-600 flex items-center justify-center text-white font-bold shrink-0 text-xs shadow-md shadow-indigo-100">
              AM
            </div>
            {!isSidebarCollapsed && (
              <div className="overflow-hidden">
                <p className="text-xs font-bold text-slate-800 truncate">Ammar ATS</p>
                <p className="text-[9px] text-slate-400 uppercase font-black tracking-tighter truncate">Inventory Manager</p>
              </div>
            )}
          </div>
        </div>
      </motion.aside>

      {/* Main Content Area */}
      <main className="flex-1 flex flex-col h-full overflow-hidden relative">
        {/* Header - Now simplified */}
        <header className="bg-white border-b border-slate-200">
          <div className="max-w-full mx-auto px-4 sm:px-6 lg:px-8 py-4">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex items-center gap-4">
                <div className="flex flex-col">
                  <h1 className="text-xl font-bold text-slate-900 leading-tight uppercase">
                    {MENU_ITEMS.find(i => i.id === activeTab)?.name || "DASHBOARD"}
                  </h1>
                  <div className="flex items-center gap-2">
                    <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                    <span className="text-[10px] uppercase font-black text-slate-400 tracking-widest">Real-time Inventory Sync</span>
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-2 flex-wrap justify-end">
                {/* Export Excel Button */}
                <button
                  onClick={activeTab === 'job-tracking' ? handleExportJobTracking : handleExportExcel}
                  className={cn(
                    "flex items-center gap-2 px-3 py-2 text-white rounded-lg transition-all text-sm font-medium shadow-sm",
                    activeTab === 'job-tracking' ? "bg-blue-600 hover:bg-blue-700" : "bg-emerald-600 hover:bg-emerald-700"
                  )}
                  title={activeTab === 'job-tracking' ? "Export Job Tracking to Excel" : "Export to Excel"}
                >
                  <Download className="w-4 h-4" />
                  {activeTab === 'job-tracking' ? "EXPORT JOBS" : "EXPORT EXCEL"}
                </button>

              {selectedMonth && (
                <div className="flex gap-2">
                  <button
                    onClick={() => {
                      const nextState = !showJobSummary;
                      setShowJobSummary(nextState);
                      if (nextState) {
                        setShowDailySummary(false);
                        setShowJobColumn(true);
                      }
                    }}
                    className={cn(
                      "flex items-center gap-2 px-3 py-2 rounded-lg transition-all text-sm font-medium shadow-sm border",
                      showJobSummary 
                        ? "bg-blue-50 text-blue-600 border-blue-200" 
                        : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                    )}
                    title="Toggle Job-wise Summary View"
                  >
                    <History className="w-4 h-4" />
                    {showJobSummary ? "SHOW PART REPORT" : "JOB SUMMARY"}
                  </button>
                  <button
                    onClick={() => {
                      setShowDailySummary(!showDailySummary);
                      if (!showDailySummary) setShowJobSummary(false);
                    }}
                    className={cn(
                      "flex items-center gap-2 px-3 py-2 rounded-lg transition-all text-sm font-medium shadow-sm border",
                      showDailySummary 
                        ? "bg-blue-50 text-blue-600 border-blue-200" 
                        : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50"
                    )}
                    title="Toggle Daily Summary View"
                  >
                    <LayoutDashboard className="w-4 h-4" />
                    {showDailySummary ? "SHOW STOCK REPORT" : "DAILY SUMMARY"}
                  </button>
                  <button
                    onClick={handleExportMonthlyDailyReport}
                    className="flex items-center gap-2 px-3 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-all text-sm font-medium shadow-sm"
                    title="Export Daily Breakdown for Selected Month"
                  >
                    <Download className="w-4 h-4" />
                    DAILY REPORT
                  </button>
                </div>
              )}

              {/* Last Updated Info */}
              {lastUpdated && (
                <div className="hidden lg:flex flex-col items-end mr-2 text-right">
                  <span className="text-[9px] text-slate-400 uppercase font-black tracking-tighter">Last Sync</span>
                  <span className="text-[11px] text-slate-500 font-mono font-bold">
                    {format(lastUpdated, 'HH:mm:ss')}
                  </span>
                </div>
              )}

              {/* Refresh Button */}
              <button 
                onClick={() => {
                  if (activeTab === 'job-tracking') {
                    fetchData('oil-seal', 0, true);
                    fetchData('bonding', 0, true);
                    fetchData('quality', 0, true);
                    fetchData('trimming', 0, true);
                  } else {
                    fetchData(activeTab, 0, true);
                  }
                }}
                className="flex items-center gap-2 px-3 py-2 bg-slate-100 hover:bg-blue-50 text-slate-600 hover:text-blue-600 rounded-lg transition-all text-sm font-medium"
                title="Refresh Data"
              >
                <RefreshCw className={cn("w-4 h-4", loading && "animate-spin")} />
                Refresh
              </button>

              {/* Reset Filters Button */}
              <button 
                onClick={resetFilters}
                className="flex items-center gap-2 px-3 py-2 bg-rose-50 hover:bg-rose-100 text-rose-600 rounded-lg transition-all text-sm font-medium border border-rose-100"
                title="Reset all filters to default"
              >
                <XCircle className="w-4 h-4" />
                Reset Filters
              </button>
            </div>
          </div>
        </div>
      </header>

      {/* Primary Sticky Scroll Area for content */}
      <div className="flex-1 overflow-y-auto overflow-x-hidden custom-scrollbar">
        {/* Sticky Filter Bar */}
        <div className="sticky top-0 z-[100] bg-white border-b border-slate-200 shadow-md py-4">
          <div className="max-w-full mx-auto px-4 sm:px-6 lg:px-8">
            {/* Filters Row */}
            <div className="flex flex-wrap items-center gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200">
            {activeTab === 'job-tracking' ? (
              <div className="flex-1 flex flex-wrap items-center gap-3">
                <div className="relative flex-1 min-w-[200px]">
                  <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                  <input
                    type="text"
                    placeholder="Search Job # or Part Name..."
                    className="w-full pl-11 pr-4 py-2.5 bg-white border border-slate-200 rounded-full text-sm outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all shadow-sm"
                    value={jobSearchTerm}
                    onChange={(e) => setJobSearchTerm(e.target.value)}
                  />
                </div>
                
                <div className="flex items-center gap-2 bg-white border border-slate-200 rounded-full px-4 py-2 shadow-sm">
                  <Calendar className="w-4 h-4 text-slate-400" />
                  <input 
                    type="date" 
                    className="text-xs font-semibold outline-none border-none bg-transparent text-slate-600 focus:text-indigo-600"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                  />
                  <span className="text-slate-200 font-light mx-1">|</span>
                  <input 
                    type="date" 
                    className="text-xs font-semibold outline-none border-none bg-transparent text-slate-600 focus:text-indigo-600"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                  />
                  {(startDate || endDate) && (
                    <button 
                      onClick={() => { setStartDate(''); setEndDate(''); }}
                      className="ml-2 p-1 hover:bg-slate-50 rounded-full text-rose-500 transition-colors"
                      title="Clear date range"
                    >
                      <XCircle className="w-4 h-4" />
                    </button>
                  )}
                </div>
              </div>
            ) : (
              <>
                {/* Excel-like Filter */}
                <div className="relative flex-1 min-w-[300px]">
                <div className="flex items-center gap-2">
                  <div
                    onClick={() => setIsFilterOpen(!isFilterOpen)}
                    onKeyDown={(e) => e.key === 'Enter' && setIsFilterOpen(!isFilterOpen)}
                    role="button"
                    tabIndex={0}
                    className={cn(
                      "flex items-center justify-between w-full px-5 py-2.5 bg-white border border-slate-200 rounded-full text-sm font-bold transition-all hover:bg-slate-50 shadow-sm cursor-pointer outline-none focus:ring-2 focus:ring-indigo-500/20",
                      selectedPartNames.length > 0 ? "text-indigo-600 border-indigo-200 ring-1 ring-indigo-100" : "text-slate-600"
                    )}
                  >
                    <div className="flex items-center gap-2 truncate">
                      <Filter className="w-4 h-4 flex-shrink-0" />
                      <span className="truncate">
                        {selectedPartNames.length === 0 
                          ? "Select Items (Excel Filter)..." 
                          : `${selectedPartNames.length} items selected`}
                      </span>
                    </div>
                    <div className="flex items-center gap-1">
                      {selectedPartNames.length > 0 && (
                        <button 
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedPartNames([]);
                          }}
                          className="p-1 hover:bg-rose-100 text-rose-600 rounded transition-colors"
                          title="Clear item selection"
                        >
                          <RefreshCw className="w-3 h-3" />
                        </button>
                      )}
                      <ChevronDown className={cn("w-4 h-4 transition-transform flex-shrink-0", isFilterOpen && "rotate-180")} />
                    </div>
                  </div>
                </div>

                {isFilterOpen && (
                  <>
                    <div 
                      className="fixed inset-0 z-40" 
                      onClick={() => setIsFilterOpen(false)} 
                    />
                    <div className="absolute top-full left-0 mt-2 w-full min-w-[320px] bg-white border border-slate-200 rounded-xl shadow-2xl z-50 overflow-hidden flex flex-col max-h-[500px]">
                      <div className="p-3 border-b border-slate-100 bg-slate-50 space-y-3">
                        <div className="flex items-center justify-between">
                          <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Filter Items</span>
                          <div className="flex gap-3">
                            <button 
                              onClick={() => {
                                // Only select those that are currently visible in the filter search
                                const newSelection = Array.from(new Set([...selectedPartNames, ...filteredPartNamesInDropdown]));
                                setSelectedPartNames(newSelection);
                              }}
                              className="text-[10px] font-bold text-blue-600 uppercase hover:underline"
                            >
                              Select All
                            </button>
                            <button 
                              onClick={() => {
                                // Only clear those that are currently visible in the filter search
                                const filteredNamesSet = new Set(filteredPartNamesInDropdown);
                                setSelectedPartNames(selectedPartNames.filter(name => !filteredNamesSet.has(name)));
                              }}
                              className="text-[10px] font-bold text-rose-600 uppercase hover:underline"
                            >
                              Clear All
                            </button>
                          </div>
                        </div>
                        <div className="relative">
                          <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                          <input
                            type="text"
                            placeholder="SEARCH ITEMS IN LIST..."
                            className="w-full pl-11 pr-11 py-2.5 bg-white border border-slate-200 rounded-full text-sm outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all shadow-sm"
                            value={filterSearch}
                            onChange={(e) => setFilterSearch(e.target.value)}
                            autoFocus
                          />
                          {filterSearch && (
                            <button
                              onClick={() => setFilterSearch('')}
                              className="absolute right-3 top-1/2 -translate-y-1/2 p-1 hover:bg-slate-100 rounded-full transition-colors"
                            >
                              <XCircle className="w-3 h-3 text-slate-400" />
                            </button>
                          )}
                        </div>
                      </div>
                      <div className="overflow-y-auto p-2 space-y-1 custom-scrollbar">
                        {filteredPartNamesInDropdown.length > 0 ? (
                          filteredPartNamesInDropdown.map(name => (
                            <label 
                              key={name} 
                              className={cn(
                                "flex items-center gap-3 p-2 hover:bg-blue-50 rounded-lg cursor-pointer transition-colors group",
                                selectedPartNames.includes(name) && "bg-blue-50/50"
                              )}
                            >
                              <input
                                type="checkbox"
                                className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                                checked={selectedPartNames.includes(name)}
                                onChange={(e) => {
                                  if (e.target.checked) {
                                    setSelectedPartNames([...selectedPartNames, name]);
                                  } else {
                                    setSelectedPartNames(selectedPartNames.filter(n => n !== name));
                                  }
                                }}
                              />
                              <span className={cn(
                                "text-sm transition-colors uppercase font-medium",
                                selectedPartNames.includes(name) ? "text-blue-700" : "text-slate-600 group-hover:text-slate-900"
                              )}>
                                {name}
                              </span>
                            </label>
                          ))
                        ) : (
                          <div className="p-4 text-center text-slate-400 text-sm italic">
                            No items match your search
                          </div>
                        )}
                      </div>
                    </div>
                  </>
                )}
              </div>

              <div className="h-8 w-px bg-slate-200 hidden md:block" />

              {/* Excel-like Date Filter */}
              <div className="relative flex-1 min-w-[250px]">
                <div className="flex items-center gap-2">
                  <div
                    onClick={() => setIsDateFilterOpen(!isDateFilterOpen)}
                    onKeyDown={(e) => e.key === 'Enter' && setIsDateFilterOpen(!isDateFilterOpen)}
                    role="button"
                    tabIndex={0}
                    className={cn(
                      "flex items-center justify-between w-full px-5 py-2.5 bg-white border border-slate-200 rounded-full text-sm font-bold transition-all hover:bg-slate-50 shadow-sm cursor-pointer outline-none focus:ring-2 focus:ring-indigo-500/20",
                      selectedDates.length > 0 ? "text-indigo-600 border-indigo-200 ring-1 ring-indigo-100" : "text-slate-600"
                    )}
                  >
                    <div className="flex items-center gap-2 truncate">
                      <Calendar className="w-4 h-4 flex-shrink-0" />
                      <span className="truncate">
                        {selectedDates.length === 0 
                          ? "Select Dates (Excel Filter)..." 
                          : `${selectedDates.length} dates selected`}
                      </span>
                    </div>
                    <div className="flex items-center gap-1">
                      {selectedDates.length > 0 && (
                        <button 
                          onClick={(e) => {
                            e.stopPropagation();
                            setSelectedDates([]);
                          }}
                          className="p-1 hover:bg-rose-100 text-rose-600 rounded transition-colors"
                          title="Clear date selection"
                        >
                          <RefreshCw className="w-3 h-3" />
                        </button>
                      )}
                      <ChevronDown className={cn("w-4 h-4 transition-transform flex-shrink-0", isDateFilterOpen && "rotate-180")} />
                    </div>
                  </div>
                </div>

                {isDateFilterOpen && (
                  <>
                    <div 
                      className="fixed inset-0 z-40" 
                      onClick={() => setIsDateFilterOpen(false)} 
                    />
                    <div className="absolute top-full left-0 mt-2 w-full min-w-[300px] bg-white border border-slate-200 rounded-xl shadow-2xl z-50 overflow-hidden flex flex-col max-h-[500px]">
                      <div className="p-3 border-b border-slate-100 bg-slate-50 space-y-3">
                        <div className="flex items-center justify-between">
                          <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">Filter Dates</span>
                          <div className="flex gap-3">
                            <button 
                              onClick={() => {
                                const newSelection = Array.from(new Set([...selectedDates, ...filteredDatesInDropdown]));
                                setSelectedDates(newSelection);
                              }}
                              className="text-[10px] font-bold text-blue-600 uppercase hover:underline"
                            >
                              Select All
                            </button>
                            <button 
                              onClick={() => {
                                const filteredDatesSet = new Set(filteredDatesInDropdown);
                                setSelectedDates(selectedDates.filter(d => !filteredDatesSet.has(d)));
                              }}
                              className="text-[10px] font-bold text-rose-600 uppercase hover:underline"
                            >
                              Clear All
                            </button>
                          </div>
                        </div>
                        <div className="relative">
                          <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                          <input
                            type="text"
                            placeholder="SEARCH DATES..."
                            className="w-full pl-11 pr-11 py-2.5 bg-white border border-slate-200 rounded-full text-sm outline-none focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 transition-all shadow-sm"
                            value={dateFilterSearch}
                            onChange={(e) => setDateFilterSearch(e.target.value)}
                            autoFocus
                          />
                          {dateFilterSearch && (
                            <button
                              onClick={() => setDateFilterSearch('')}
                              className="absolute right-3 top-1/2 -translate-y-1/2 p-1 hover:bg-slate-100 rounded-full transition-colors"
                            >
                              <XCircle className="w-3 h-3 text-slate-400" />
                            </button>
                          )}
                        </div>
                      </div>
                      <div className="overflow-y-auto p-2 space-y-1 custom-scrollbar">
                        {filteredDatesInDropdown.length > 0 ? (
                          filteredDatesInDropdown.map(date => (
                            <label 
                              key={date} 
                              className={cn(
                                "flex items-center gap-3 p-2 hover:bg-blue-50 rounded-lg cursor-pointer transition-colors group",
                                selectedDates.includes(date) && "bg-blue-50/50"
                              )}
                            >
                              <input
                                type="checkbox"
                                className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                                checked={selectedDates.includes(date)}
                                onChange={(e) => {
                                  if (e.target.checked) {
                                    setSelectedDates([...selectedDates, date]);
                                  } else {
                                    setSelectedDates(selectedDates.filter(d => d !== date));
                                  }
                                }}
                              />
                              <span className={cn(
                                "text-sm transition-colors",
                                selectedDates.includes(date) ? "text-blue-700 font-medium" : "text-slate-600 group-hover:text-slate-900"
                              )}>
                                {date}
                              </span>
                            </label>
                          ))
                        ) : (
                          <div className="p-4 text-center text-slate-400 text-sm italic">
                            No dates found
                          </div>
                        )}
                      </div>
                    </div>
                  </>
                )}
              </div>

              <div className="h-8 w-px bg-slate-200 hidden md:block" />
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-400 uppercase">Month:</span>
                <div className="relative">
                  <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                  <select
                    className="pl-10 pr-8 py-2 bg-white border border-slate-200 focus:ring-2 focus:ring-blue-500 rounded-lg text-sm appearance-none outline-none cursor-pointer transition-all"
                    value={selectedMonth}
                    onChange={(e) => setSelectedMonth(e.target.value)}
                  >
                    {availableMonths.map(month => (
                      <option key={month} value={month}>{month}</option>
                    ))}
                  </select>
                  <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400 pointer-events-none" />
                </div>
              </div>

              <div className="h-8 w-px bg-slate-200 hidden md:block" />

              {/* Date Range Filter */}
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-400 uppercase">Range:</span>
                <div className="flex items-center gap-2">
                  <input
                    type="date"
                    className="px-3 py-2 bg-white border border-slate-200 focus:ring-2 focus:ring-blue-500 rounded-lg text-sm outline-none transition-all"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                  />
                  <span className="text-slate-400 text-xs">to</span>
                  <input
                    type="date"
                    className="px-3 py-2 bg-white border border-slate-200 focus:ring-2 focus:ring-blue-500 rounded-lg text-sm outline-none transition-all"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                  />
                  {(startDate || endDate) && (
                    <button 
                      onClick={() => { setStartDate(''); setEndDate(''); }}
                      className="text-xs text-blue-600 hover:text-blue-700 font-medium"
                    >
                      Clear
                    </button>
                  )}
                </div>
              </div>
            </>
          )}
          </div>
        </div>
      </div>

      <div className="max-w-full mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {error ? (
          <div className="bg-red-50 border border-red-200 rounded-xl p-6 flex items-start gap-4">
            <AlertCircle className="w-6 h-6 text-red-600 flex-shrink-0 mt-0.5" />
            <div>
              <h3 className="text-red-900 font-semibold mb-1">Data Fetching Error</h3>
              <p className="text-red-700 text-sm">{error}</p>
              <button 
                onClick={() => {
                  if (activeTab === 'job-tracking') {
                    fetchData('oil-seal', 0, true);
                    fetchData('bonding', 0, true);
                    fetchData('quality', 0, true);
                    fetchData('trimming', 0, true);
                  } else {
                    fetchData(activeTab, 0, true);
                  }
                }}
                className="mt-3 text-sm font-medium text-red-900 hover:underline"
              >
                Try again
              </button>
            </div>
          </div>
        ) : (
          <div className="space-y-6">
            {activeTab === 'job-tracking' ? (
              <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
                <div className="p-6 border-b border-slate-100 bg-slate-50/50 flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div className="flex items-center gap-4">
                      <div className="bg-blue-600 p-3 rounded-xl shadow-lg shadow-blue-200">
                        <Search className="w-6 h-6 text-white" />
                      </div>
                      <div>
                        <h2 className="text-xl font-black text-slate-800 tracking-tight uppercase">JOB NUMBER TRACKING</h2>
                        <p className="text-sm text-slate-500 font-medium uppercase">TRACK PIECE MOVEMENT ACROSS ALL DEPARTMENTS BY JOB #</p>
                      </div>
                    </div>
                  <div className="flex items-center gap-4">
                    <button
                      onClick={() => setShowDuplicatesOnly(!showDuplicatesOnly)}
                      className={cn(
                        "flex items-center gap-2 px-4 py-2.5 rounded-xl text-sm font-black transition-all border shadow-sm",
                        showDuplicatesOnly 
                          ? "bg-amber-600 border-amber-600 text-white shadow-amber-200" 
                          : "bg-white border-slate-200 text-slate-600 hover:bg-slate-50"
                      )}
                    >
                      <AlertTriangle className={cn("w-4 h-4", showDuplicatesOnly ? "text-white" : "text-amber-500")} />
                      {showDuplicatesOnly ? "SHOWING DUPLICATES" : `SHOW DUPLICATES (${duplicateCount})`}
                    </button>
                    <div className="flex items-center gap-6 bg-white px-6 py-3 rounded-2xl border border-slate-100 shadow-sm">
                      <div className="text-center">
                        <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Total Jobs</span>
                        <span className="text-2xl font-black text-blue-600 leading-none">{jobTrackingData.length}</span>
                      </div>
                      <div className="h-8 w-px bg-slate-100" />
                      <div className="text-center">
                        <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest block mb-1">Total Pcs</span>
                        <span className="text-2xl font-black text-emerald-600 leading-none">
                          {jobTrackingData.reduce((sum, job) => sum + Object.values(job.totals as Record<string, number>).reduce((s, v) => s + v, 0), 0).toLocaleString()}
                        </span>
                      </div>
                    </div>
                  </div>
                </div>
                
                <div className="overflow-x-auto custom-scrollbar">
                  <table className="w-full border-collapse table-fixed">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-200">
                        <th 
                          style={{ width: `${jobColWidth}px` }}
                          className="px-6 py-4 text-left text-[11px] font-black text-slate-500 uppercase tracking-widest border-r border-slate-200 relative group/header"
                        >
                          Job #
                          <div 
                            onMouseDown={(e) => handleResizeMouseDown(e, 'job')}
                            className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-blue-400/50 transition-colors z-10"
                            title="Drag to resize"
                          />
                        </th>
                        <th 
                          style={{ width: `${partColWidth}px` }}
                          className="px-6 py-4 text-left text-[11px] font-black text-slate-500 uppercase tracking-widest border-r border-slate-200 relative group/header"
                        >
                          Part Name
                          <div 
                            onMouseDown={(e) => handleResizeMouseDown(e, 'part')}
                            className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-blue-400/50 transition-colors z-10"
                            title="Drag to resize"
                          />
                        </th>
                        <th className="px-6 py-4 text-left text-[11px] font-black text-slate-500 uppercase tracking-widest">Transaction Summary (Quantity by Type)</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {jobTrackingData.length > 0 ? (
                        (showAllJobs ? jobTrackingData : jobTrackingData.slice(0, 10)).map((job) => (
                          <React.Fragment key={job.uniqueKey}>
                            <tr 
                              onClick={() => setExpandedJob(expandedJob === job.uniqueKey ? null : job.uniqueKey)}
                              className={cn(
                                "hover:bg-blue-50/30 transition-all group cursor-pointer",
                                expandedJob === job.uniqueKey && "bg-blue-50/50"
                              )}
                            >
                              <td style={{ width: `${jobColWidth}px` }} className="px-6 py-5 border-r border-slate-100 truncate">
                                <div className="flex items-center gap-3">
                                  <div className={cn(
                                    "p-1 rounded-md transition-colors",
                                    expandedJob === job.uniqueKey ? "bg-blue-600 text-white" : "bg-slate-100 text-slate-400 group-hover:bg-blue-100 group-hover:text-blue-600"
                                  )}>
                                    {expandedJob === job.uniqueKey ? <ChevronDown className="w-3 h-3" /> : <ChevronRight className="w-3 h-3" />}
                                  </div>
                                  <div className="flex flex-col">
                                    <div className="flex items-center gap-2">
                                      <span className="text-sm font-black text-blue-700 tracking-tight uppercase">{job.jobId}</span>
                                      {job.isDuplicate && (
                                        <div 
                                          className="flex items-center gap-1 px-1.5 py-0.5 bg-amber-100 text-amber-700 rounded text-[9px] font-black uppercase tracking-tighter animate-pulse"
                                          title={`Duplicate Job #: This Job ID is used for multiple parts: ${job.duplicatePartNames.join(', ')}`}
                                        >
                                          <AlertTriangle className="w-2.5 h-2.5" />
                                          Duplicate
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                </div>
                              </td>
                              <td style={{ width: `${partColWidth}px` }} className="px-6 py-5 border-r border-slate-100 truncate" title={job.partName?.toUpperCase()}>
                                <span className="text-sm font-bold text-slate-700 uppercase">{job.partName}</span>
                              </td>
                              <td className="px-6 py-5">
                                <div className="flex flex-wrap gap-2">
                                  {job.sortedTotalsEntries
                                    .slice(0, 10).map(([compositeKey, qty]) => {
                                      const [dept, type] = compositeKey.split('|');
                                      return (
                                        <div key={compositeKey} className="flex items-center gap-2 pl-2 pr-3 py-1.5 bg-white border border-slate-200 rounded-xl shadow-sm hover:border-blue-200 hover:shadow-md transition-all">
                                          <div className={cn(
                                            "w-1.5 h-1.5 rounded-full",
                                            type.toLowerCase().includes('in') ? "bg-emerald-500" : "bg-rose-500"
                                          )} />
                                          <div className="flex flex-col">
                                            <span className="text-[11px] font-black text-black uppercase tracking-tighter leading-none mb-0.5">{dept.replace('-', ' ').toUpperCase()}</span>
                                            <span className="text-[10px] font-black text-slate-500 uppercase tracking-tighter leading-none">{type.toUpperCase()}</span>
                                          </div>
                                          <span className={cn(
                                            "text-xs font-black ml-1",
                                            type.toLowerCase().includes('in') ? "text-emerald-600" : "text-rose-600"
                                          )}>
                                            {qty.toLocaleString()}
                                          </span>
                                        </div>
                                      );
                                    })}
                                  {Object.keys(job.totals).length > 10 && (
                                    <div className="flex items-center px-3 py-1.5 bg-slate-50 border border-slate-200 rounded-xl text-[10px] font-black text-slate-400 uppercase tracking-widest">
                                      +{Object.keys(job.totals).length - 10} More Types
                                    </div>
                                  )}
                                </div>
                              </td>
                            </tr>
                            {expandedJob === job.uniqueKey && (
                              <tr>
                                <td colSpan={3} className="px-6 py-6 bg-slate-50/80 border-b border-slate-200">
                                  <div className="space-y-4">
                                    <div className="flex items-center justify-between">
                                      <h4 className="text-[11px] font-black text-slate-500 uppercase tracking-widest flex items-center gap-2">
                                        <History className="w-3.5 h-3.5" />
                                        DETAILED TRANSACTION HISTORY
                                      </h4>
                                      <span className="text-[10px] font-bold text-slate-400 uppercase">
                                        {job.transactions.length} TOTAL MOVEMENTS RECORDED
                                      </span>
                                    </div>
                                    <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-2">
                                      {job.transactions.map((t: any, i: number) => (
                                        <div key={i} className="bg-white p-2 rounded-xl border border-slate-200 shadow-sm flex flex-col gap-1.5 hover:shadow-md transition-all border-l-4 border-l-blue-500">
                                          <div className="flex justify-between items-start">
                                            <div className="flex flex-col">
                                              <span className="text-xs font-black text-black uppercase tracking-tight leading-none mb-1">{t.department?.toUpperCase()}</span>
                                              <span className="text-[10px] font-bold text-slate-800 uppercase">{t.type?.toUpperCase()}</span>
                                            </div>
                                            <div className="bg-slate-50 px-1.5 py-0.5 rounded text-[9px] font-mono font-bold text-slate-500">
                                              {t.date}
                                            </div>
                                          </div>
                                          <div className="flex justify-between items-end">
                                            <div className="text-right w-full">
                                              <span className="text-[8px] font-black text-slate-400 uppercase tracking-tighter leading-none block mb-0.5">Quantity</span>
                                              <span className={cn(
                                                "text-xs font-black",
                                                t.type.toLowerCase().includes('in') ? "text-emerald-600" : "text-rose-600"
                                              )}>
                                                {t.quantity.toLocaleString()}
                                              </span>
                                            </div>
                                          </div>
                                        </div>
                                      ))}
                                    </div>
                                  </div>
                                </td>
                              </tr>
                            )}
                          </React.Fragment>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={3} className="px-6 py-24 text-center">
                            <div className="flex flex-col items-center gap-4">
                              <div className="bg-slate-50 p-6 rounded-full">
                                <Search className="w-12 h-12 text-slate-300" />
                              </div>
                              <div className="max-w-xs">
                                <p className="text-slate-500 font-bold text-lg mb-1 uppercase">
                                  {jobSearchTerm ? "NO JOBS FOUND" : "READY TO TRACK"}
                                </p>
                                <p className="text-slate-400 text-sm uppercase">
                                  {jobSearchTerm 
                                    ? `WE COULDN'T FIND ANY TRANSACTIONS FOR "${jobSearchTerm.toUpperCase()}". TRY A DIFFERENT JOB #.` 
                                    : "ENTER A JOB # OR PART NAME IN THE SEARCH BAR ABOVE TO SEE ITS COMPLETE TRANSACTION HISTORY."}
                                </p>
                              </div>
                            </div>
                          </td>
                        </tr>
                      )}
                      {!showAllJobs && jobTrackingData.length > 10 && (
                        <tr>
                          <td colSpan={3} className="px-6 py-6 text-center bg-slate-50/30 border-t border-slate-100">
                            <button 
                              onClick={() => setShowAllJobs(true)}
                              className="inline-flex items-center gap-2 px-6 py-3 bg-white border border-slate-200 rounded-xl shadow-sm text-sm font-black text-blue-600 hover:bg-blue-50 hover:border-blue-200 transition-all"
                            >
                              <History className="w-4 h-4" />
                              VIEW ALL {jobTrackingData.length} JOBS
                              <ChevronDown className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <React.Fragment>
                {/* Stats Overview */}
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                  {activeTab !== 'molding' && (
                    <StatCard 
                      label="TOTAL OPENING STOCK" 
                      value={summaryData.reduce((sum, item) => sum + item.openingStock, 0).toLocaleString(undefined, { maximumFractionDigits: 0 })} 
                      icon={<Package className="w-5 h-5" />}
                      color="blue"
                    />
                  )}
                  {activeTab !== 'molding' && (
                    <StatCard 
                      label="TOTAL IN" 
                      value={summaryData.reduce((sum, item) => sum + item.totalIn, 0).toLocaleString(undefined, { maximumFractionDigits: 0 })} 
                      icon={<ArrowUpCircle className="w-5 h-5" />}
                      color="emerald"
                    />
                  )}
                  <StatCard 
                    label="TOTAL OUT" 
                    value={summaryData.reduce((sum, item) => sum + item.totalOut, 0).toLocaleString(undefined, { maximumFractionDigits: 0 })} 
                    icon={<ArrowDownCircle className="w-5 h-5" />}
                    color="rose"
                  />
                  {activeTab !== 'molding' && (
                    <StatCard 
                      label="CURRENT STOCK" 
                      value={summaryData.reduce((sum, item) => sum + item.currentStock, 0).toLocaleString(undefined, { maximumFractionDigits: 0 })} 
                      icon={<AlertCircle className="w-5 h-5" />}
                      color="amber"
                    />
                  )}
                </div>

            {/* Main Table or Daily Summary */}
            {showDailySummary ? (
              <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                <div className="p-4 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
                  <div className="flex items-center gap-4">
                    <h3 className="font-bold text-slate-800 flex items-center gap-2">
                      <Calendar className="w-5 h-5 text-blue-600" />
                      Daily Transaction Summary - {selectedMonth}
                    </h3>
                    <button
                      onClick={handleExportDailySummaryTable}
                      className="flex items-center gap-2 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg transition-all text-xs font-bold shadow-sm"
                      title="Export this summary to Excel"
                    >
                      <Download className="w-3.5 h-3.5" />
                      Export Summary
                    </button>
                  </div>
                  <span className="text-xs text-slate-500 font-medium">Showing totals for each date with transactions</span>
                </div>
                <div className="overflow-x-auto scrollbar-thin">
                  <table className="w-full text-left border-collapse table-fixed min-w-max">
                    <thead>
                      {/* Totals Row at Top */}
                      <tr className="bg-slate-50 border-b border-slate-200">
                        <th className="px-6 py-3 text-xs font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 w-[200px]">TOTALS</th>
                      { (activeTab === 'bonding') ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.metalStoreIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.chemicalStoreIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.phosphateIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.moldIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 bg-emerald-50/30 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.injcMoldOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.oilSealOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.hvcmOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.rejectionOutToMetalStore?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        ) : activeTab === 'auto-clave' ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.autoClaveProdIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.autoClaveMiniStoreIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.autoClaveMetalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.autoClaveReworkIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 bg-emerald-50/30 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.autoClaveRejectionOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.autoClaveMetalOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        ) : activeTab === 'phosphate' ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.metalStoreIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.chemicalStoreIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 bg-emerald-50/30 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.phosphateOutToBonding?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.rejectionOutToRps?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        ) : activeTab === 'oil-seal' ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.moldIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.reworkIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 bg-emerald-50/30 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.qcOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.rejectionOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        ) : activeTab === 'extrusion' ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.reworkIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.extrusionProdIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.extrusionMetalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.extrusionMiniStoreIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 bg-emerald-50/30 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.rejectionOutToRps?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.fgOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.extrusionTrimOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.qcOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        ) : activeTab === 'molding' ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.rejectionOutToRps?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.oilSealTrimmingOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        ) : activeTab === 'mini-store' ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.openingStock?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.compoundIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.moldReturnIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-blue-700 text-center bg-blue-50/30 border-r border-slate-200 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.vendorOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.injectOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.oilSealOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.moldOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.extrusionOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.autoClaveOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.labOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.rejectionOutToRps?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        ) : activeTab === 'fg-store' ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.openingStock?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.customerRejectionIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.qcIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.reworkIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.autoClaveIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 bg-emerald-50/30 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.rejectionOutToRps?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.qcReworkOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.fgOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 border-r border-slate-200 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-blue-700 text-center bg-blue-50/30 w-[120px]">{dailyTotals.currentStock?.toLocaleString()}</th>
                          </>
                        ) : activeTab === 'quality' ? (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.fgReworkIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.metalStoreIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.customerRejectionIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.oilSealTrimmingIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.extrusionIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 bg-emerald-50/30 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.rejectionOutToRps?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.metalStoreOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.oilSealTrimmingOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.fgOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.extrusionOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        ) : (
                          <>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingVendorIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingQcReworkIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingMoldingIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingMetalStoreIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingExtrusionIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 bg-emerald-50/30 w-[120px]">{dailyTotals.totalIn?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingQcOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingVendorOut?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{dailyTotals.trimmingRejectionOutToRps?.toLocaleString()}</th>
                            <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center bg-rose-50/30 w-[120px]">{dailyTotals.totalOut?.toLocaleString()}</th>
                          </>
                        )}
                      </tr>
                      {/* Headers Row with Sorting */}
                      <tr className="bg-white border-b border-slate-200 sticky top-0 z-20">
                        <th 
                          onClick={() => handleDailySort('date')}
                          className="px-6 py-4 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[200px]"
                        >
                          <div className="flex items-center justify-center gap-2">
                            <span>DATE</span>
                            <DailySortIcon field="date" />
                          </div>
                        </th>
                      { (activeTab === 'bonding') ? (
                          <>
                            <th onClick={() => handleDailySort('metalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE IN</span><DailySortIcon field="metalStoreIn" /></div></th>
                            <th onClick={() => handleDailySort('chemicalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CHEMICAL STORE IN</span><DailySortIcon field="chemicalStoreIn" /></div></th>
                            <th onClick={() => handleDailySort('phosphateIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>PHOSPHATE IN</span><DailySortIcon field="phosphateIn" /></div></th>
                            <th onClick={() => handleDailySort('moldIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MOLD IN</span><DailySortIcon field="moldIn" /></div></th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div></th>
                            <th onClick={() => handleDailySort('injcMoldOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>INJC MOLD OUT</span><DailySortIcon field="injcMoldOut" /></div></th>
                            <th onClick={() => handleDailySort('oilSealOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OIL SEAL OUT</span><DailySortIcon field="oilSealOut" /></div></th>
                            <th onClick={() => handleDailySort('hvcmOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>HVCM OUT</span><DailySortIcon field="hvcmOut" /></div></th>
                            <th onClick={() => handleDailySort('rejectionOutToMetalStore')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT</span><DailySortIcon field="rejectionOutToMetalStore" /></div></th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div></th>
                          </>
                        ) : activeTab === 'auto-clave' ? (
                          <>
                            <th onClick={() => handleDailySort('autoClaveProdIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>PROD IN</span><DailySortIcon field="autoClaveProdIn" /></div></th>
                            <th onClick={() => handleDailySort('autoClaveMiniStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MINI STORE IN</span><DailySortIcon field="autoClaveMiniStoreIn" /></div></th>
                            <th onClick={() => handleDailySort('autoClaveMetalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL IN</span><DailySortIcon field="autoClaveMetalIn" /></div></th>
                            <th onClick={() => handleDailySort('autoClaveReworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REWORK IN</span><DailySortIcon field="autoClaveReworkIn" /></div></th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div></th>
                            <th onClick={() => handleDailySort('autoClaveRejectionOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT</span><DailySortIcon field="autoClaveRejectionOut" /></div></th>
                            <th onClick={() => handleDailySort('autoClaveMetalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL OUT</span><DailySortIcon field="autoClaveMetalOut" /></div></th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div></th>
                          </>
                        ) : activeTab === 'phosphate' ? (
                          <>
                            <th onClick={() => handleDailySort('metalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE IN</span><DailySortIcon field="metalStoreIn" /></div></th>
                            <th onClick={() => handleDailySort('chemicalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CHEMICAL STORE IN</span><DailySortIcon field="chemicalStoreIn" /></div></th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div></th>
                            <th onClick={() => handleDailySort('phosphateOutToBonding')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>PHOSPHATE OUT TO BONDING</span><DailySortIcon field="phosphateOutToBonding" /></div></th>
                            <th onClick={() => handleDailySort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><DailySortIcon field="rejectionOutToRps" /></div></th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div></th>
                          </>
                        ) : activeTab === 'oil-seal' ? (
                          <>
                            <th onClick={() => handleDailySort('moldIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>MOLD IN</span><DailySortIcon field="moldIn" /></div>
                            </th>
                            <th onClick={() => handleDailySort('reworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>REWORK IN</span><DailySortIcon field="reworkIn" /></div>
                            </th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div>
                            </th>
                            <th onClick={() => handleDailySort('qcOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>QC OUT</span><DailySortIcon field="qcOut" /></div>
                            </th>
                            <th onClick={() => handleDailySort('rejectionOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>REJECTION OUT</span><DailySortIcon field="rejectionOut" /></div>
                            </th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div>
                            </th>
                          </>
                        ) : activeTab === 'extrusion' ? (
                          <>
                            <th onClick={() => handleDailySort('reworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>REWORK IN</span><DailySortIcon field="reworkIn" /></div>
                            </th>
                            <th onClick={() => handleDailySort('extrusionProdIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>PROD IN</span><DailySortIcon field="extrusionProdIn" /></div>
                            </th>
                            <th onClick={() => handleDailySort('extrusionMetalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>METAL IN EXTRUS</span><DailySortIcon field="extrusionMetalIn" /></div>
                            </th>
                            <th onClick={() => handleDailySort('extrusionMiniStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>MINI STORE IN EXTRUS</span><DailySortIcon field="extrusionMiniStoreIn" /></div>
                            </th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div>
                            </th>
                            <th onClick={() => handleDailySort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><DailySortIcon field="rejectionOutToRps" /></div>
                            </th>
                            <th onClick={() => handleDailySort('fgOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>FG OUT</span><DailySortIcon field="fgOut" /></div>
                            </th>
                            <th onClick={() => handleDailySort('extrusionTrimOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>TRIM OUT</span><DailySortIcon field="extrusionTrimOut" /></div>
                            </th>
                            <th onClick={() => handleDailySort('qcOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>QC OUT</span><DailySortIcon field="qcOut" /></div>
                            </th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div>
                            </th>
                          </>
                        ) : activeTab === 'molding' ? (
                          <>
                            <th onClick={() => handleDailySort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><DailySortIcon field="rejectionOutToRps" /></div>
                            </th>
                            <th onClick={() => handleDailySort('oilSealTrimmingOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>OIL SEAL TRIMMING OUT</span><DailySortIcon field="oilSealTrimmingOut" /></div>
                            </th>
                            <th onClick={() => handleDailySort('trimmingOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>TRIMMING OUT</span><DailySortIcon field="trimmingOut" /></div>
                            </th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div>
                            </th>
                          </>
                        ) : activeTab === 'mini-store' ? (
                          <>
                            <th onClick={() => handleDailySort('openingStock')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OPENING STOCK</span><DailySortIcon field="openingStock" /></div></th>
                            <th onClick={() => handleDailySort('compoundIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>COMPOUND IN</span><DailySortIcon field="compoundIn" /></div></th>
                            <th onClick={() => handleDailySort('moldReturnIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MOLD RETURN IN</span><DailySortIcon field="moldReturnIn" /></div></th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div></th>
                            <th onClick={() => handleDailySort('vendorOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>VENDOR OUT</span><DailySortIcon field="vendorOut" /></div></th>
                            <th onClick={() => handleDailySort('injectOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>INJECT OUT</span><DailySortIcon field="injectOut" /></div></th>
                            <th onClick={() => handleDailySort('oilSealOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OIL SEAL OUT</span><DailySortIcon field="oilSealOut" /></div></th>
                            <th onClick={() => handleDailySort('moldOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MOLD OUT</span><DailySortIcon field="moldOut" /></div></th>
                            <th onClick={() => handleDailySort('extrusionOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>EXTRUSION OUT</span><DailySortIcon field="extrusionOut" /></div></th>
                            <th onClick={() => handleDailySort('autoClaveOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>AUTOCLAVE OUT</span><DailySortIcon field="autoClaveOut" /></div></th>
                            <th onClick={() => handleDailySort('labOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>LAB OUT</span><DailySortIcon field="labOut" /></div></th>
                            <th onClick={() => handleDailySort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><DailySortIcon field="rejectionOutToRps" /></div></th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div></th>
                          </>
                        ) : activeTab === 'fg-store' ? (
                          <>
                            <th onClick={() => handleDailySort('openingStock')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OPENING STOCK</span><DailySortIcon field="openingStock" /></div></th>
                            <th onClick={() => handleDailySort('customerRejectionIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CUSTOMER REJECTION IN</span><DailySortIcon field="customerRejectionIn" /></div></th>
                            <th onClick={() => handleDailySort('qcIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC IN</span><DailySortIcon field="qcIn" /></div></th>
                            <th onClick={() => handleDailySort('reworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REWORK IN</span><DailySortIcon field="reworkIn" /></div></th>
                            <th onClick={() => handleDailySort('autoClaveIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>AUTO CLAVE IN</span><DailySortIcon field="autoClaveIn" /></div></th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div></th>
                            <th onClick={() => handleDailySort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><DailySortIcon field="rejectionOutToRps" /></div></th>
                            <th onClick={() => handleDailySort('qcReworkOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC REWORK OUT</span><DailySortIcon field="qcReworkOut" /></div></th>
                            <th onClick={() => handleDailySort('fgOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>FG OUT</span><DailySortIcon field="fgOut" /></div></th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div></th>
                            <th onClick={() => handleDailySort('currentStock')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CURRENT STOCK</span><DailySortIcon field="currentStock" /></div></th>
                          </>
                        ) : activeTab === 'quality' ? (
                          <>
                            <th onClick={() => handleDailySort('fgReworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>FG REWORK IN</span><DailySortIcon field="fgReworkIn" /></div></th>
                            <th onClick={() => handleDailySort('metalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE IN</span><DailySortIcon field="metalStoreIn" /></div></th>
                            <th onClick={() => handleDailySort('customerRejectionIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CUSTOMER REJECTION IN</span><DailySortIcon field="customerRejectionIn" /></div></th>
                            <th onClick={() => handleDailySort('oilSealTrimmingIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OIL SEAL TRIMMING IN</span><DailySortIcon field="oilSealTrimmingIn" /></div></th>
                            <th onClick={() => handleDailySort('trimmingIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TRIMMING IN</span><DailySortIcon field="trimmingIn" /></div></th>
                            <th onClick={() => handleDailySort('extrusionIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>EXTRUSION IN</span><DailySortIcon field="extrusionIn" /></div></th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div></th>
                            <th onClick={() => handleDailySort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><DailySortIcon field="rejectionOutToRps" /></div></th>
                            <th onClick={() => handleDailySort('metalStoreOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE OUT</span><DailySortIcon field="metalStoreOut" /></div></th>
                            <th onClick={() => handleDailySort('oilSealTrimmingOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OIL SEAL TRIMMING OUT</span><DailySortIcon field="oilSealTrimmingOut" /></div></th>
                            <th onClick={() => handleDailySort('trimmingOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TRIMMING OUT</span><DailySortIcon field="trimmingOut" /></div></th>
                            <th onClick={() => handleDailySort('fgOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>FG OUT</span><DailySortIcon field="fgOut" /></div></th>
                            <th onClick={() => handleDailySort('extrusionOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>EXTRUSION OUT</span><DailySortIcon field="extrusionOut" /></div></th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div></th>
                          </>
                        ) : (
                          <>
                            <th onClick={() => handleDailySort('trimmingVendorIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>VENDOR IN</span><DailySortIcon field="trimmingVendorIn" /></div></th>
                            <th onClick={() => handleDailySort('trimmingQcReworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC REWORK IN</span><DailySortIcon field="trimmingQcReworkIn" /></div></th>
                            <th onClick={() => handleDailySort('trimmingMoldingIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MOLD IN</span><DailySortIcon field="trimmingMoldingIn" /></div></th>
                            <th onClick={() => handleDailySort('trimmingMetalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE IN</span><DailySortIcon field="trimmingMetalStoreIn" /></div></th>
                            <th onClick={() => handleDailySort('trimmingExtrusionIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>EXTRUSION IN</span><DailySortIcon field="trimmingExtrusionIn" /></div></th>
                            <th onClick={() => handleDailySort('totalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL IN</span><DailySortIcon field="totalIn" /></div></th>
                            <th onClick={() => handleDailySort('trimmingQcOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC OUT</span><DailySortIcon field="trimmingQcOut" /></div></th>
                            <th onClick={() => handleDailySort('trimmingVendorOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>VENDOR OUT</span><DailySortIcon field="trimmingVendorOut" /></div></th>
                            <th onClick={() => handleDailySort('trimmingRejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECT TO RPS</span><DailySortIcon field="trimmingRejectionOutToRps" /></div></th>
                            <th onClick={() => handleDailySort('totalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TOTAL OUT</span><DailySortIcon field="totalOut" /></div></th>
                          </>
                        )}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {sortedDailySummary.length > 0 ? (
                        sortedDailySummary.map((day, idx) => (
                          <tr key={idx} className="hover:bg-slate-50 transition-colors">
                            <td className="px-6 py-4 text-sm font-bold text-slate-900 text-center border-r border-slate-100 bg-slate-50/30">
                              {format(day.date, 'dd-MMM-yyyy')}
                            </td>
                            { (activeTab === 'bonding') ? (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.metalStoreIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.chemicalStoreIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.phosphateIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.moldIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-emerald-600 text-center border-r border-slate-100 bg-emerald-50/20">{day.totalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.injcMoldOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.oilSealOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.hvcmOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.rejectionOutToMetalStore?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                            ) : activeTab === 'auto-clave' ? (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.autoClaveProdIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.autoClaveMiniStoreIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.autoClaveMetalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.autoClaveReworkIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-emerald-600 text-center border-r border-slate-100 bg-emerald-50/20">{day.totalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.autoClaveRejectionOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.autoClaveMetalOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                            ) : activeTab === 'phosphate' ? (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.metalStoreIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.chemicalStoreIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-emerald-600 text-center border-r border-slate-100 bg-emerald-50/20">{day.totalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.phosphateOutToBonding?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.rejectionOutToRps?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                            ) : activeTab === 'oil-seal' ? (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.moldIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.reworkIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-emerald-600 text-center border-r border-slate-100 bg-emerald-50/20">{day.totalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.qcOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.rejectionOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                            ) : activeTab === 'extrusion' ? (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.reworkIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.extrusionProdIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.extrusionMetalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.extrusionMiniStoreIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-emerald-600 text-center border-r border-slate-100 bg-emerald-50/20">{day.totalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.rejectionOutToRps?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.fgOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.extrusionTrimOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.qcOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                            ) : activeTab === 'molding' ? (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.rejectionOutToRps?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.oilSealTrimmingOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                            ) : activeTab === 'mini-store' ? (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.openingStock?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.compoundIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.moldReturnIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-blue-600 text-center bg-blue-50/20 border-r border-slate-100">{day.totalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.vendorOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.injectOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.oilSealOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.moldOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.extrusionOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.autoClaveOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.labOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.rejectionOutToRps?.toLocaleString() || 0}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                          ) : activeTab === 'fg-store' ? (
                            <>
                              <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.openingStock?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.customerRejectionIn?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.qcIn?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.reworkIn?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.autoClaveIn?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] font-bold text-emerald-600 text-center border-r border-slate-100 bg-emerald-50/20">{day.totalIn?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.rejectionOutToRps?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.qcReworkOut?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.fgOut?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20 border-r border-slate-100">{day.totalOut?.toLocaleString()}</td>
                              <td className="px-2 py-2 text-[11px] font-bold text-blue-600 text-center bg-blue-50/20">{day.currentStock?.toLocaleString()}</td>
                            </>
                          ) : activeTab === 'quality' ? (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.fgReworkIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.metalStoreIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.customerRejectionIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.oilSealTrimmingIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.extrusionIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-emerald-600 text-center border-r border-slate-100 bg-emerald-50/20">{day.totalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.rejectionOutToRps?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.metalStoreOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.oilSealTrimmingOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.fgOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.extrusionOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                            ) : (
                              <>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingVendorIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingQcReworkIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingMoldingIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingMetalStoreIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingExtrusionIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-emerald-600 text-center border-r border-slate-100 bg-emerald-50/20">{day.totalIn?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingQcOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingVendorOut?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] text-slate-600 text-center border-r border-slate-100">{day.trimmingRejectionOutToRps?.toLocaleString()}</td>
                                <td className="px-2 py-2 text-[11px] font-bold text-rose-600 text-center bg-rose-50/20">{day.totalOut?.toLocaleString()}</td>
                              </>
                            )}
                          </tr>
                        ))
                      ) : (
                        <tr>
                          <td colSpan={20} className="px-6 py-12 text-center text-slate-400 italic">
                            No daily transaction data found for this month.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            ) : (
              <React.Fragment>
                {/* Top Scrollbar for small screens with many headers */}
                <div 
                  ref={topScrollRef} 
                  className="overflow-x-auto h-2 mb-1 scrollbar-thin scrollbar-thumb-slate-300 scrollbar-track-transparent"
                  style={{ width: '100%' }}
                >
                  <div style={{ height: '1px' }}></div>
                </div>

                {/* Main Table */}
                <div className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                  {showJobSummary && (
                    <div className="p-4 border-b border-slate-100 bg-slate-50 flex items-center justify-between">
                      <div className="flex items-center gap-4">
                        <h3 className="font-bold text-slate-800 flex items-center gap-2">
                          <History className="w-5 h-5 text-blue-600" />
                          Job Summary - {selectedMonth}
                        </h3>
                        <button
                          onClick={handleExportExcel}
                          className="flex items-center gap-2 px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg transition-all text-xs font-bold shadow-sm"
                          title="Export Job Summary to Excel"
                        >
                          <Download className="w-3.5 h-3.5" />
                          Export Job Summary
                        </button>
                      </div>
                      <span className="text-xs text-slate-500 font-medium whitespace-nowrap">Showing aggregated totals per Job #</span>
                    </div>
                  )}
                  <div ref={tableScrollRef} className="overflow-x-auto scrollbar-thin">
                    <table className="w-full text-left border-collapse min-w-max table-fixed">
                  <thead>
                        {/* Totals Row */}
                        <tr className="bg-slate-50 border-b border-slate-200">
                          <th 
                            style={{ width: `${mainPartColWidth}px` }}
                            className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider sticky left-0 bg-slate-50 z-30 border-r border-slate-200"
                          >
                            TOTALS
                          </th>
                          {showJobColumn && (
                            <th 
                              style={{ width: `${mainJobColWidth}px` }}
                              className="px-2 py-2 text-[11px] font-black text-slate-900 text-center border-r border-slate-200"
                            >
                              {/* Empty for Job # */}
                            </th>
                          )}
                          {activeTab !== 'molding' && (
                            <th className="px-2 py-2 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">
                              {columnTotals.openingStock.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                            </th>
                          )}
                          {activeTab === 'trimming' && (
                            <th className="px-2 py-2 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">
                              {columnTotals.vendorOpeningStock.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                            </th>
                          )}
                          
                          { (activeTab === 'bonding') ? (
                            <>
                              {isColVisible(columnTotals.metalStoreIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.metalStoreIn.toLocaleString()}</th>}
                              {isColVisible(columnTotals.chemicalStoreIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.chemicalStoreIn.toLocaleString()}</th>}
                              {isColVisible(columnTotals.phosphateIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.phosphateIn.toLocaleString()}</th>}
                              {isColVisible(columnTotals.moldIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.moldIn.toLocaleString()}</th>}
                            </>
                          ) : activeTab === 'auto-clave' ? (
                            <>
                              {isColVisible(columnTotals.autoClaveProdIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.autoClaveProdIn.toLocaleString()}</th>}
                              {isColVisible(columnTotals.autoClaveMiniStoreIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.autoClaveMiniStoreIn.toLocaleString()}</th>}
                              {isColVisible(columnTotals.autoClaveMetalIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.autoClaveMetalIn.toLocaleString()}</th>}
                              {isColVisible(columnTotals.autoClaveReworkIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.autoClaveReworkIn.toLocaleString()}</th>}
                            </>
                          ) : activeTab === 'phosphate' ? (
                            <>
                              {isColVisible(columnTotals.metalStoreIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.metalStoreIn.toLocaleString()}</th>}
                              {isColVisible(columnTotals.chemicalStoreIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.chemicalStoreIn.toLocaleString()}</th>}
                            </>
                          ) : activeTab === 'oil-seal' ? (
                            <>
                              {isColVisible(columnTotals.moldIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.moldIn.toLocaleString()}</th>}
                              {isColVisible(columnTotals.reworkIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.reworkIn.toLocaleString()}</th>}
                            </>
                          ) : activeTab === 'extrusion' ? (
                            <>
                              <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.reworkIn?.toLocaleString()}</th>
                              <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.extrusionProdIn?.toLocaleString()}</th>
                              <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.extrusionMetalIn?.toLocaleString()}</th>
                              <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.extrusionMiniStoreIn?.toLocaleString()}</th>
                            </>
                          ) : activeTab === 'molding' ? (
                        <>
                          {/* No IN columns for molding */}
                        </>
                      ) : activeTab === 'quality' ? (
    <>
      {isColVisible(columnTotals.fgReworkIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.fgReworkIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.metalStoreIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.metalStoreIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.customerRejectionIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.customerRejectionIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.oilSealTrimmingIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.oilSealTrimmingIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.trimmingIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.extrusionIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.extrusionIn?.toLocaleString()}</th>}
    </>
  ) : activeTab === 'mini-store' ? (
    <>
      {isColVisible(columnTotals.compoundIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.compoundIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.moldReturnIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.moldReturnIn?.toLocaleString()}</th>}
    </>
  ) : activeTab === 'fg-store' ? (
    <>
      {isColVisible(columnTotals.customerRejectionIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.customerRejectionIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.qcIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.qcIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.reworkIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.reworkIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.autoClaveIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.autoClaveIn?.toLocaleString()}</th>}
    </>
  ) : (
    // Trimming
    <>
      {isColVisible(columnTotals.trimmingVendorIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingVendorIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.trimmingQcReworkIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingQcReworkIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.trimmingMoldingIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingMoldingIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.trimmingMetalStoreIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingMetalStoreIn?.toLocaleString()}</th>}
      {isColVisible(columnTotals.trimmingExtrusionIn) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingExtrusionIn?.toLocaleString()}</th>}
    </>
                      )}

                          {activeTab !== 'molding' && (
                            <th className="px-2 py-1.5 text-[11px] font-black text-emerald-700 text-center border-r border-slate-200 w-[120px]">
                              {columnTotals.totalIn.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                            </th>
                          )}

                          { (activeTab === 'bonding') ? (
                            <>
                              {isColVisible(columnTotals.injcMoldOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.injcMoldOut?.toLocaleString()}</th>}
                              {isColVisible(columnTotals.oilSealOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.oilSealOut?.toLocaleString()}</th>}
                              {isColVisible(columnTotals.hvcmOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.hvcmOut?.toLocaleString()}</th>}
                              {isColVisible(columnTotals.rejectionOutToMetalStore) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.rejectionOutToMetalStore?.toLocaleString()}</th>}
                            </>
                          ) : activeTab === 'auto-clave' ? (
                            <>
                              {isColVisible(columnTotals.autoClaveRejectionOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.autoClaveRejectionOut?.toLocaleString()}</th>}
                              {isColVisible(columnTotals.autoClaveMetalOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.autoClaveMetalOut?.toLocaleString()}</th>}
                            </>
                          ) : activeTab === 'phosphate' ? (
                            <>
                              {isColVisible(columnTotals.phosphateOutToBonding) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.phosphateOutToBonding?.toLocaleString()}</th>}
                              {isColVisible(columnTotals.rejectionOutToRps) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.rejectionOutToRps?.toLocaleString()}</th>}
                            </>
                          ) : activeTab === 'oil-seal' ? (
                            <>
                              {isColVisible(columnTotals.qcOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.qcOut?.toLocaleString()}</th>}
                              {isColVisible(columnTotals.rejectionOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.rejectionOut?.toLocaleString()}</th>}
                            </>
                          ) : activeTab === 'extrusion' ? (
                            <>
                              <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.rejectionOutToRps?.toLocaleString()}</th>
                              <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.fgOut?.toLocaleString()}</th>
                              <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.extrusionTrimOut?.toLocaleString()}</th>
                              <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.qcOut?.toLocaleString()}</th>
                            </>
                          ) : activeTab === 'molding' ? (
                        <>
                          {isColVisible(columnTotals.rejectionOutToRps) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.rejectionOutToRps?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.oilSealTrimmingOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.oilSealTrimmingOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.trimmingOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingOut?.toLocaleString()}</th>}
                        </>
                      ) : activeTab === 'quality' ? (
                        <>
                          {isColVisible(columnTotals.rejectionOutToRps) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.rejectionOutToRps?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.metalStoreOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.metalStoreOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.oilSealTrimmingOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.oilSealTrimmingOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.trimmingOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.fgOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.fgOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.extrusionOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.extrusionOut?.toLocaleString()}</th>}
                        </>
                      ) : activeTab === 'mini-store' ? (
                        <>
                          {isColVisible(columnTotals.vendorOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.vendorOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.injectOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.injectOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.oilSealOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.oilSealOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.moldOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.moldOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.extrusionOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.extrusionOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.autoClaveOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.autoClaveOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.labOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.labOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.rejectionOutToRps) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.rejectionOutToRps?.toLocaleString()}</th>}
                        </>
                      ) : activeTab === 'fg-store' ? (
                        <>
                          {isColVisible(columnTotals.rejectionOutToRps) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.rejectionOutToRps?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.qcReworkOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.qcReworkOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.fgOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.fgOut?.toLocaleString()}</th>}
                        </>
                      ) : (
                        // Trimming
                        <>
                          {isColVisible(columnTotals.trimmingQcOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingQcOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.trimmingVendorOut) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingVendorOut?.toLocaleString()}</th>}
                          {isColVisible(columnTotals.trimmingRejectionOutToRps) && <th className="px-2 py-1.5 text-[11px] font-black text-slate-900 text-center border-r border-slate-200 w-[120px]">{columnTotals.trimmingRejectionOutToRps?.toLocaleString()}</th>}
                        </>
                      )}

                      <th className="px-2 py-1.5 text-[11px] font-black text-rose-700 text-center border-r border-slate-200 w-[120px]">
                        {columnTotals.totalOut.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                      </th>
                      {activeTab !== 'molding' && (
                        <th className={cn(
                          "px-2 py-1.5 text-[11px] font-black text-blue-700 text-center w-[120px]",
                          (hasAnyNextMonthStock || activeTab === 'trimming') && "border-r border-slate-200"
                        )}>
                          {columnTotals.currentStock.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                        </th>
                      )}
                      {activeTab === 'trimming' && (
                        <>
                          <th className="px-2 py-1.5 text-[11px] font-black text-blue-700 text-center border-r border-slate-200 w-[120px]">
                            {columnTotals.vendorStock.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                          </th>
                          <th className={cn(
                            "px-2 py-1.5 text-[11px] font-black text-blue-700 text-center w-[120px]",
                            hasAnyNextMonthStock && "border-r border-slate-200"
                          )}>
                            {columnTotals.totalStock.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                          </th>
                        </>
                      )}
                      {activeTab !== 'molding' && hasAnyNextMonthStock && (
                        <th className="px-2 py-1.5 text-[11px] font-black text-indigo-700 text-center w-[120px]">
                          {columnTotals.nextMonthOpeningStock?.toLocaleString(undefined, { maximumFractionDigits: 0 })}
                        </th>
                      )}
                    </tr>
                    {/* Labels Row */}
                    <tr className="bg-white border-b border-slate-200 sticky top-0 z-20">
                      <th 
                        onClick={() => handleSort('itemId')}
                        style={{ width: `${mainPartColWidth}px` }}
                        className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider sticky left-0 bg-white z-30 border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors relative group/header"
                      >
                        <div className="flex items-center justify-between gap-2">
                          <div className="flex items-center gap-2">
                            <span>{showJobSummary ? (activeTab === 'mini-store' ? 'JOB #' : 'JOB #') : (activeTab === 'mini-store' ? 'ITEM ID' : 'PART NO. & NAME')}</span>
                            <div className="flex items-center gap-1">
                              {!showJobSummary && (
                                <button 
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setIsFilterOpen(!isFilterOpen);
                                  }}
                                  className={cn(
                                    "p-0.5 rounded hover:bg-slate-200 transition-colors",
                                    selectedPartNames.length > 0 ? "text-blue-600 bg-blue-50" : "text-black"
                                  )}
                                  title="Filter Part Names"
                                >
                                  <Filter className="w-2.5 h-2.5" />
                                </button>
                              )}
                              <button 
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setShowJobColumn(!showJobColumn);
                                }}
                                className={cn(
                                  "p-0.5 rounded hover:bg-slate-200 transition-colors",
                                  !showJobColumn ? "text-blue-600 bg-blue-50" : "text-black"
                                )}
                                title={showJobColumn ? "Hide Job # Column" : "Show Job # Column"}
                              >
                                {showJobColumn ? <EyeOff className="w-2.5 h-2.5" /> : <Eye className="w-2.5 h-2.5" />}
                              </button>
                            </div>
                          </div>
                          <SortIcon field="itemId" />
                        </div>
                        <div 
                          onMouseDown={(e) => handleResizeMouseDown(e, 'mainPart')}
                          className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-blue-400/50 transition-colors z-40"
                          title="Drag to resize"
                        />
                      </th>
                      {showJobColumn && (
                        <th 
                          onClick={() => handleSort('jobId')}
                          style={{ width: `${mainJobColWidth}px` }}
                          className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors relative group/header"
                        >
                          <div className="flex items-center justify-center gap-2">
                            <span>{showJobSummary ? (activeTab === 'mini-store' ? 'ITEM ID' : 'PART NAME') : (activeTab === 'mini-store' ? 'JOB #' : 'JOB #')}</span>
                            <SortIcon field="jobId" />
                          </div>
                          <div 
                            onMouseDown={(e) => handleResizeMouseDown(e, 'mainJob')}
                            className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-blue-400/50 transition-colors z-10"
                            title="Drag to resize"
                          />
                        </th>
                      )}
                      {activeTab !== 'molding' && (
                        <th 
                          onClick={() => handleSort('openingStock')}
                          className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"
                        >
                          <div className="flex items-center justify-center gap-2">
                            <span>{activeTab === 'trimming' || activeTab === 'molding' ? 'OPENING STOCK' : (activeTab === 'mini-store' ? 'OPENING STOCK' : 'TOTAL OPENING STOCK')}</span>
                            <SortIcon field="openingStock" />
                          </div>
                        </th>
                      )}
                      {activeTab === 'trimming' && (
                        <th 
                          onClick={() => handleSort('vendorOpeningStock')}
                          className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"
                        >
                          <div className="flex items-center justify-center gap-2">
                            <span>VENDOR OPENING STOCK</span>
                            <SortIcon field="vendorOpeningStock" />
                          </div>
                        </th>
                      )}
                      
                      { (activeTab === 'bonding') ? (
                        <>
                          {isColVisible(columnTotals.metalStoreIn) && <th onClick={() => handleSort('metalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE IN</span><SortIcon field="metalStoreIn" /></div></th>}
                          {isColVisible(columnTotals.chemicalStoreIn) && <th onClick={() => handleSort('chemicalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CHEMICAL STORE IN</span><SortIcon field="chemicalStoreIn" /></div></th>}
                          {isColVisible(columnTotals.phosphateIn) && <th onClick={() => handleSort('phosphateIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>PHOSPHATE IN</span><SortIcon field="phosphateIn" /></div></th>}
                          {isColVisible(columnTotals.moldIn) && <th onClick={() => handleSort('moldIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MOLD IN</span><SortIcon field="moldIn" /></div></th>}
                        </>
                      ) : activeTab === 'auto-clave' ? (
                        <>
                          {isColVisible(columnTotals.autoClaveProdIn) && <th onClick={() => handleSort('autoClaveProdIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>PROD IN</span><SortIcon field="autoClaveProdIn" /></div></th>}
                          {isColVisible(columnTotals.autoClaveMiniStoreIn) && <th onClick={() => handleSort('autoClaveMiniStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MINI STORE IN</span><SortIcon field="autoClaveMiniStoreIn" /></div></th>}
                          {isColVisible(columnTotals.autoClaveMetalIn) && <th onClick={() => handleSort('autoClaveMetalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL IN</span><SortIcon field="autoClaveMetalIn" /></div></th>}
                          {isColVisible(columnTotals.autoClaveReworkIn) && <th onClick={() => handleSort('autoClaveReworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REWORK IN</span><SortIcon field="autoClaveReworkIn" /></div></th>}
                        </>
                      ) : activeTab === 'phosphate' ? (
                        <>
                          {isColVisible(columnTotals.metalStoreIn) && <th onClick={() => handleSort('metalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE IN</span><SortIcon field="metalStoreIn" /></div></th>}
                          {isColVisible(columnTotals.chemicalStoreIn) && <th onClick={() => handleSort('chemicalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CHEMICAL STORE IN</span><SortIcon field="chemicalStoreIn" /></div></th>}
                        </>
                      ) : activeTab === 'oil-seal' ? (
                        <>
                          {isColVisible(columnTotals.moldIn) && (
                            <th onClick={() => handleSort('moldIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>MOLD IN</span><SortIcon field="moldIn" /></div>
                            </th>
                          )}
                          {isColVisible(columnTotals.reworkIn) && (
                            <th onClick={() => handleSort('reworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>REWORK IN</span><SortIcon field="reworkIn" /></div>
                            </th>
                          )}
                        </>
                      ) : activeTab === 'extrusion' ? (
                        <>
                          <th onClick={() => handleSort('reworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REWORK IN</span><SortIcon field="reworkIn" /></div></th>
                          <th onClick={() => handleSort('extrusionProdIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>PROD IN</span><SortIcon field="extrusionProdIn" /></div></th>
                          <th onClick={() => handleSort('extrusionMetalIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL IN EXTRUS</span><SortIcon field="extrusionMetalIn" /></div></th>
                          <th onClick={() => handleSort('extrusionMiniStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MINI STORE IN EXTRUS</span><SortIcon field="extrusionMiniStoreIn" /></div></th>
                        </>
                      ) : activeTab === 'molding' ? (
                        <>
                          {/* No IN columns for molding */}
                        </>
                      ) : activeTab === 'quality' ? (
                        <>
                          {isColVisible(columnTotals.fgReworkIn) && <th onClick={() => handleSort('fgReworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>FG REWORK IN</span><SortIcon field="fgReworkIn" /></div></th>}
                          {isColVisible(columnTotals.metalStoreIn) && <th onClick={() => handleSort('metalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE IN</span><SortIcon field="metalStoreIn" /></div></th>}
                          {isColVisible(columnTotals.customerRejectionIn) && <th onClick={() => handleSort('customerRejectionIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CUSTOMER REJECTION IN</span><SortIcon field="customerRejectionIn" /></div></th>}
                          {isColVisible(columnTotals.oilSealTrimmingIn) && <th onClick={() => handleSort('oilSealTrimmingIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OIL SEAL TRIMMING IN</span><SortIcon field="oilSealTrimmingIn" /></div></th>}
                          {isColVisible(columnTotals.trimmingIn) && <th onClick={() => handleSort('trimmingIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TRIMMING IN</span><SortIcon field="trimmingIn" /></div></th>}
                          {isColVisible(columnTotals.extrusionIn) && <th onClick={() => handleSort('extrusionIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>EXTRUSION IN</span><SortIcon field="extrusionIn" /></div></th>}
                        </>
                      ) : activeTab === 'mini-store' ? (
                        <>
                          {isColVisible(columnTotals.compoundIn) && <th onClick={() => handleSort('compoundIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>COMPOUND IN</span><SortIcon field="compoundIn" /></div></th>}
                          {isColVisible(columnTotals.moldReturnIn) && <th onClick={() => handleSort('moldReturnIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MOLD RETURN IN</span><SortIcon field="moldReturnIn" /></div></th>}
                        </>
                      ) : activeTab === 'fg-store' ? (
                        <>
                          {isColVisible(columnTotals.customerRejectionIn) && <th onClick={() => handleSort('customerRejectionIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>CUSTOMER REJECTION IN</span><SortIcon field="customerRejectionIn" /></div></th>}
                          {isColVisible(columnTotals.qcIn) && <th onClick={() => handleSort('qcIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC IN</span><SortIcon field="qcIn" /></div></th>}
                          {isColVisible(columnTotals.reworkIn) && <th onClick={() => handleSort('reworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REWORK IN</span><SortIcon field="reworkIn" /></div></th>}
                          {isColVisible(columnTotals.autoClaveIn) && <th onClick={() => handleSort('autoClaveIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>AUTO CLAVE IN</span><SortIcon field="autoClaveIn" /></div></th>}
                        </>
                      ) : (
                        // Trimming
                        <>
                          {isColVisible(columnTotals.trimmingVendorIn) && <th onClick={() => handleSort('trimmingVendorIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TRIMMING VENDOR IN</span><SortIcon field="trimmingVendorIn" /></div></th>}
                          {isColVisible(columnTotals.trimmingQcReworkIn) && <th onClick={() => handleSort('trimmingQcReworkIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC REWORK IN</span><SortIcon field="trimmingQcReworkIn" /></div></th>}
                          {isColVisible(columnTotals.trimmingMoldingIn) && <th onClick={() => handleSort('trimmingMoldingIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MOLD IN</span><SortIcon field="trimmingMoldingIn" /></div></th>}
                          {isColVisible(columnTotals.trimmingMetalStoreIn) && <th onClick={() => handleSort('trimmingMetalStoreIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE IN</span><SortIcon field="trimmingMetalStoreIn" /></div></th>}
                          {isColVisible(columnTotals.trimmingExtrusionIn) && <th onClick={() => handleSort('trimmingExtrusionIn')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>EXTRUSION IN</span><SortIcon field="trimmingExtrusionIn" /></div></th>}
                        </>
                      )}

                      {activeTab !== 'molding' && (
                        <th 
                          onClick={() => handleSort('totalIn')}
                          className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"
                        >
                          <div className="flex items-center justify-center gap-2">
                            <span>TOTAL IN</span>
                            <SortIcon field="totalIn" />
                          </div>
                        </th>
                      )}

                      { (activeTab === 'bonding') ? (
                        <>
                          {isColVisible(columnTotals.injcMoldOut) && <th onClick={() => handleSort('injcMoldOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>INJC MOLD OUT</span><SortIcon field="injcMoldOut" /></div></th>}
                          {isColVisible(columnTotals.oilSealOut) && <th onClick={() => handleSort('oilSealOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OIL SEAL OUT</span><SortIcon field="oilSealOut" /></div></th>}
                          {isColVisible(columnTotals.hvcmOut) && <th onClick={() => handleSort('hvcmOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>HVCM OUT</span><SortIcon field="hvcmOut" /></div></th>}
                          {isColVisible(columnTotals.rejectionOutToMetalStore) && <th onClick={() => handleSort('rejectionOutToMetalStore')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT</span><SortIcon field="rejectionOutToMetalStore" /></div></th>}
                        </>
                      ) : activeTab === 'auto-clave' ? (
                        <>
                          {isColVisible(columnTotals.autoClaveRejectionOut) && <th onClick={() => handleSort('autoClaveRejectionOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT</span><SortIcon field="autoClaveRejectionOut" /></div></th>}
                          {isColVisible(columnTotals.autoClaveMetalOut) && <th onClick={() => handleSort('autoClaveMetalOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL OUT</span><SortIcon field="autoClaveMetalOut" /></div></th>}
                        </>
                      ) : activeTab === 'phosphate' ? (
                        <>
                          {isColVisible(columnTotals.phosphateOutToBonding) && <th onClick={() => handleSort('phosphateOutToBonding')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>PHOSPHATE OUT</span><SortIcon field="phosphateOutToBonding" /></div></th>}
                          {isColVisible(columnTotals.rejectionOutToRps) && <th onClick={() => handleSort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><SortIcon field="rejectionOutToRps" /></div></th>}
                        </>
                      ) : activeTab === 'oil-seal' ? (
                        <>
                          {isColVisible(columnTotals.qcOut) && (
                            <th onClick={() => handleSort('qcOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>QC OUT</span><SortIcon field="qcOut" /></div>
                            </th>
                          )}
                          {isColVisible(columnTotals.rejectionOut) && (
                            <th onClick={() => handleSort('rejectionOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>REJECTION OUT</span><SortIcon field="rejectionOut" /></div>
                            </th>
                          )}
                        </>
                      ) : activeTab === 'extrusion' ? (
                        <>
                          <th onClick={() => handleSort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><SortIcon field="rejectionOutToRps" /></div></th>
                          <th onClick={() => handleSort('fgOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>FG OUT</span><SortIcon field="fgOut" /></div></th>
                          <th onClick={() => handleSort('extrusionTrimOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TRIM OUT</span><SortIcon field="extrusionTrimOut" /></div></th>
                          <th onClick={() => handleSort('qcOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC OUT</span><SortIcon field="qcOut" /></div></th>
                        </>
                      ) : activeTab === 'molding' ? (
                        <>
                          {isColVisible(columnTotals.rejectionOutToRps) && (
                            <th onClick={() => handleSort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><SortIcon field="rejectionOutToRps" /></div>
                            </th>
                          )}
                          {isColVisible(columnTotals.oilSealTrimmingOut) && (
                            <th onClick={() => handleSort('oilSealTrimmingOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>OIL SEAL TRIMMING OUT</span><SortIcon field="oilSealTrimmingOut" /></div>
                            </th>
                          )}
                          {isColVisible(columnTotals.trimmingOut) && (
                            <th onClick={() => handleSort('trimmingOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]">
                              <div className="flex items-center justify-center gap-2"><span>TRIMMING OUT</span><SortIcon field="trimmingOut" /></div>
                            </th>
                          )}
                        </>
                      ) : activeTab === 'quality' ? (
                        <>
                          {isColVisible(columnTotals.rejectionOutToRps) && <th onClick={() => handleSort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><SortIcon field="rejectionOutToRps" /></div></th>}
                          {isColVisible(columnTotals.metalStoreOut) && <th onClick={() => handleSort('metalStoreOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>METAL STORE OUT</span><SortIcon field="metalStoreOut" /></div></th>}
                          {isColVisible(columnTotals.oilSealTrimmingOut) && <th onClick={() => handleSort('oilSealTrimmingOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OIL SEAL TRIMMING OUT</span><SortIcon field="oilSealTrimmingOut" /></div></th>}
                          {isColVisible(columnTotals.trimmingOut) && <th onClick={() => handleSort('trimmingOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>TRIMMING OUT</span><SortIcon field="trimmingOut" /></div></th>}
                          {isColVisible(columnTotals.fgOut) && <th onClick={() => handleSort('fgOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>FG OUT</span><SortIcon field="fgOut" /></div></th>}
                          {isColVisible(columnTotals.extrusionOut) && <th onClick={() => handleSort('extrusionOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>EXTRUSION OUT</span><SortIcon field="extrusionOut" /></div></th>}
                        </>
                      ) : activeTab === 'mini-store' ? (
                        <>
                          {isColVisible(columnTotals.vendorOut) && <th onClick={() => handleSort('vendorOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>VENDOR OUT</span><SortIcon field="vendorOut" /></div></th>}
                          {isColVisible(columnTotals.injectOut) && <th onClick={() => handleSort('injectOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>INJECT OUT</span><SortIcon field="injectOut" /></div></th>}
                          {isColVisible(columnTotals.oilSealOut) && <th onClick={() => handleSort('oilSealOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>OIL SEAL OUT</span><SortIcon field="oilSealOut" /></div></th>}
                          {isColVisible(columnTotals.moldOut) && <th onClick={() => handleSort('moldOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>MOLD OUT</span><SortIcon field="moldOut" /></div></th>}
                          {isColVisible(columnTotals.extrusionOut) && <th onClick={() => handleSort('extrusionOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>EXTRUSION OUT</span><SortIcon field="extrusionOut" /></div></th>}
                          {isColVisible(columnTotals.autoClaveOut) && <th onClick={() => handleSort('autoClaveOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>AUTOCLAVE OUT</span><SortIcon field="autoClaveOut" /></div></th>}
                          {isColVisible(columnTotals.labOut) && <th onClick={() => handleSort('labOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>LAB OUT</span><SortIcon field="labOut" /></div></th>}
                          {isColVisible(columnTotals.rejectionOutToRps) && <th onClick={() => handleSort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><SortIcon field="rejectionOutToRps" /></div></th>}
                        </>
                      ) : activeTab === 'fg-store' ? (
                        <>
                          {isColVisible(columnTotals.rejectionOutToRps) && <th onClick={() => handleSort('rejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><SortIcon field="rejectionOutToRps" /></div></th>}
                          {isColVisible(columnTotals.qcReworkOut) && <th onClick={() => handleSort('qcReworkOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC REWORK OUT</span><SortIcon field="qcReworkOut" /></div></th>}
                          {isColVisible(columnTotals.fgOut) && <th onClick={() => handleSort('fgOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>FG OUT</span><SortIcon field="fgOut" /></div></th>}
                        </>
                      ) : (
                        // Trimming
                        <>
                          {isColVisible(columnTotals.trimmingQcOut) && <th onClick={() => handleSort('trimmingQcOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>QC OUT</span><SortIcon field="trimmingQcOut" /></div></th>}
                          {isColVisible(columnTotals.trimmingVendorOut) && <th onClick={() => handleSort('trimmingVendorOut')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>VENDOR OUT</span><SortIcon field="trimmingVendorOut" /></div></th>}
                          {isColVisible(columnTotals.trimmingRejectionOutToRps) && <th onClick={() => handleSort('trimmingRejectionOutToRps')} className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"><div className="flex items-center justify-center gap-2"><span>REJECTION OUT TO RPS</span><SortIcon field="trimmingRejectionOutToRps" /></div></th>}
                        </>
                      )}

                      <th 
                        onClick={() => handleSort('totalOut')}
                        className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"
                      >
                        <div className="flex items-center justify-center gap-2">
                          <span>TOTAL OUT</span>
                          <SortIcon field="totalOut" />
                        </div>
                      </th>
                      {activeTab !== 'molding' && (
                        <th 
                          onClick={() => handleSort('currentStock')}
                          className={cn(
                            "px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]",
                            (hasAnyNextMonthStock || activeTab === 'trimming') && "border-r border-slate-200"
                          )}
                        >
                          <div className="flex items-center justify-center gap-2">
                            <span>{activeTab === 'trimming' ? 'IN HOUSE STOCK' : (activeTab === 'mini-store' ? 'TOTAL STOCK' : 'CURRENT STOCK')}</span>
                            <SortIcon field="currentStock" />
                          </div>
                        </th>
                      )}
                      {activeTab === 'trimming' && (
                        <>
                          <th 
                            onClick={() => handleSort('vendorStock')}
                            className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center border-r border-slate-200 cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"
                          >
                            <div className="flex items-center justify-center gap-2">
                              <span>VENDOR STOCK</span>
                              <SortIcon field="vendorStock" />
                            </div>
                          </th>
                          <th 
                            onClick={() => handleSort('totalStock')}
                            className={cn(
                              "px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]",
                              hasAnyNextMonthStock && "border-r border-slate-200"
                            )}
                          >
                            <div className="flex items-center justify-center gap-2">
                              <span>TOTAL STOCK</span>
                              <SortIcon field="totalStock" />
                            </div>
                          </th>
                        </>
                      )}
                      {activeTab !== 'molding' && hasAnyNextMonthStock && (
                        <th 
                          onClick={() => handleSort('nextMonthOpeningStock')}
                          className="px-2 py-2 text-[10px] font-bold text-black uppercase tracking-wider text-center cursor-pointer hover:bg-slate-50 transition-colors w-[120px]"
                        >
                          <div className="flex items-center justify-center gap-2">
                            <span>PHYSICAL STOCK</span>
                            <SortIcon field="nextMonthOpeningStock" />
                          </div>
                        </th>
                      )}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {summaryData.length > 0 ? (
                      summaryData.map((item, idx) => (
                        <tr key={idx} className="hover:bg-slate-50 transition-colors group">
                          <td 
                            style={{ width: `${mainPartColWidth}px` }}
                            className="px-2 py-1.5 text-[11px] font-medium text-slate-700 sticky left-0 bg-white group-hover:bg-slate-50 z-10 border-r border-slate-200 truncate" 
                            title={item.itemId?.toUpperCase()}
                          >
                            {item.itemId?.toUpperCase()}
                          </td>
                          {showJobColumn && (
                            <td 
                              style={{ width: `${mainJobColWidth}px` }}
                              className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200 truncate uppercase" 
                              title={showJobSummary ? (item as any).partName?.toUpperCase() : item.jobId?.toUpperCase()}
                            >
                              {showJobSummary ? (item as any).partName?.toUpperCase() : item.jobId?.toUpperCase()}
                            </td>
                          )}
                          {activeTab !== 'molding' && (
                            <td className="px-2 py-1.5 text-[11px] text-center font-semibold text-slate-900 bg-yellow-50/30 border-r border-slate-200">
                              {item.openingStock?.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 })}
                            </td>
                          )}
                          {activeTab === 'trimming' && (
                            <td className="px-2 py-1.5 text-[11px] text-center font-semibold text-slate-900 bg-yellow-50/30 border-r border-slate-200">
                              {item.vendorOpeningStock?.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 })}
                            </td>
                          )}
                          
                          { (activeTab === 'bonding') ? (
                            <>
                              {isColVisible(columnTotals.metalStoreIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.metalStoreIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.chemicalStoreIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.chemicalStoreIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.phosphateIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.phosphateIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.moldIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.moldIn?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'auto-clave' ? (
                            <>
                              {isColVisible(columnTotals.autoClaveProdIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.autoClaveProdIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.autoClaveMiniStoreIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.autoClaveMiniStoreIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.autoClaveMetalIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.autoClaveMetalIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.autoClaveReworkIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.autoClaveReworkIn?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'phosphate' ? (
                            <>
                              {isColVisible(columnTotals.metalStoreIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.metalStoreIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.chemicalStoreIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.chemicalStoreIn?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'oil-seal' ? (
                            <>
                              {isColVisible(columnTotals.moldIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.moldIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.reworkIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.reworkIn?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'extrusion' ? (
                            <>
                              <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.reworkIn?.toLocaleString()}</td>
                              <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.extrusionProdIn?.toLocaleString()}</td>
                              <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.extrusionMetalIn?.toLocaleString()}</td>
                              <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.extrusionMiniStoreIn?.toLocaleString()}</td>
                            </>
                          ) : activeTab === 'molding' ? (
                            <>
                              {/* No IN columns for molding */}
                            </>
                          ) : activeTab === 'quality' ? (
                            <>
                              {isColVisible(columnTotals.fgReworkIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.fgReworkIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.metalStoreIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.metalStoreIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.customerRejectionIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.customerRejectionIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.oilSealTrimmingIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.oilSealTrimmingIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.trimmingIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.extrusionIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.extrusionIn?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'mini-store' ? (
                            <>
                              {isColVisible(columnTotals.compoundIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.compoundIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.moldReturnIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.moldReturnIn?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'fg-store' ? (
                            <>
                              {isColVisible(columnTotals.customerRejectionIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.customerRejectionIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.qcIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.qcIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.reworkIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.reworkIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.autoClaveIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.autoClaveIn?.toLocaleString()}</td>}
                            </>
                          ) : (
                            // Trimming
                            <>
                              {isColVisible(columnTotals.trimmingVendorIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingVendorIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.trimmingQcReworkIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingQcReworkIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.trimmingMoldingIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingMoldingIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.trimmingMetalStoreIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingMetalStoreIn?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.trimmingExtrusionIn) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingExtrusionIn?.toLocaleString()}</td>}
                            </>
                          )}

                          {activeTab !== 'molding' && (
                            <td className="px-2 py-1.5 text-[11px] text-center font-bold text-emerald-700 bg-emerald-50/30 border-r border-slate-200">
                              {item.totalIn?.toLocaleString()}
                            </td>
                          )}

                          { (activeTab === 'bonding') ? (
                            <>
                              {isColVisible(columnTotals.injcMoldOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.injcMoldOut?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.oilSealOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.oilSealOut?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.hvcmOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.hvcmOut?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.rejectionOutToMetalStore) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.rejectionOutToMetalStore?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'auto-clave' ? (
                            <>
                              {isColVisible(columnTotals.autoClaveRejectionOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.autoClaveRejectionOut?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.autoClaveMetalOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.autoClaveMetalOut?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'phosphate' ? (
                            <>
                              {isColVisible(columnTotals.phosphateOutToBonding) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.phosphateOutToBonding?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.rejectionOutToRps) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.rejectionOutToRps?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'oil-seal' ? (
                            <>
                              {isColVisible(columnTotals.qcOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.qcOut?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.rejectionOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.rejectionOut?.toLocaleString()}</td>}
                            </>
                          ) : activeTab === 'extrusion' ? (
                            <>
                              <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.rejectionOutToRps?.toLocaleString()}</td>
                              <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.fgOut?.toLocaleString()}</td>
                              <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.extrusionTrimOut?.toLocaleString()}</td>
                              <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.qcOut?.toLocaleString()}</td>
                            </>
                          ) : activeTab === 'molding' ? (
                            <>
                              {isColVisible(columnTotals.rejectionOutToRps) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.rejectionOutToRps?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.oilSealTrimmingOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.oilSealTrimmingOut?.toLocaleString()}</td>}
                              {isColVisible(columnTotals.trimmingOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingOut?.toLocaleString()}</td>}
                            </>
                                  ) : activeTab === 'quality' ? (
                                    <>
                                      {isColVisible(columnTotals.rejectionOutToRps) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.rejectionOutToRps?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.metalStoreOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.metalStoreOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.oilSealTrimmingOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.oilSealTrimmingOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.trimmingOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.fgOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.fgOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.extrusionOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.extrusionOut?.toLocaleString()}</td>}
                                    </>
                                  ) : activeTab === 'mini-store' ? (
                                    <>
                                      {isColVisible(columnTotals.vendorOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.vendorOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.injectOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.injectOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.oilSealOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.oilSealOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.moldOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.moldOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.extrusionOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.extrusionOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.autoClaveOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.autoClaveOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.labOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.labOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.rejectionOutToRps) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.rejectionOutToRps?.toLocaleString()}</td>}
                                    </>
                                  ) : activeTab === 'fg-store' ? (
                                    <>
                                      {isColVisible(columnTotals.rejectionOutToRps) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.rejectionOutToRps?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.qcReworkOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.qcReworkOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.fgOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.fgOut?.toLocaleString()}</td>}
                                    </>
                                  ) : (
                                    // Trimming
                                    <>
                                      {isColVisible(columnTotals.trimmingQcOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingQcOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.trimmingVendorOut) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingVendorOut?.toLocaleString()}</td>}
                                      {isColVisible(columnTotals.trimmingRejectionOutToRps) && <td className="px-2 py-1.5 text-[11px] text-center text-slate-600 border-r border-slate-200">{item.trimmingRejectionOutToRps?.toLocaleString()}</td>}
                                    </>
                          )}

                          <td className="px-2 py-1.5 text-[11px] text-center font-bold text-rose-700 bg-rose-50/30 border-r border-slate-200">
                            {item.totalOut?.toLocaleString()}
                          </td>
                          {activeTab !== 'molding' && (
                            <td className={cn(
                              "px-2 py-1.5 text-[11px] text-center font-black bg-blue-50/30",
                              item.currentStock < 0 ? "text-red-600" : "text-blue-700",
                              (hasAnyNextMonthStock || activeTab === 'trimming') && "border-r border-slate-200"
                            )}>
                              {item.currentStock?.toLocaleString()}
                            </td>
                          )}
                          {activeTab === 'trimming' && (
                            <>
                              <td className={cn(
                                "px-2 py-1.5 text-[11px] text-center font-black bg-blue-50/30 border-r border-slate-200",
                                item.vendorStock < 0 ? "text-red-600" : "text-blue-700"
                              )}>
                                {item.vendorStock?.toLocaleString()}
                              </td>
                              <td className={cn(
                                "px-2 py-1.5 text-[11px] text-center font-black bg-blue-50/30",
                                item.totalStock < 0 ? "text-red-600" : "text-blue-700",
                                hasAnyNextMonthStock && "border-r border-slate-200"
                              )}>
                                {item.totalStock?.toLocaleString()}
                              </td>
                            </>
                          )}
                          {activeTab !== 'molding' && hasAnyNextMonthStock && (
                            <td className="px-2 py-1.5 text-[11px] text-center font-black text-indigo-700 bg-indigo-50/30">
                              {item.nextMonthOpeningStock?.toLocaleString() ?? '-'}
                            </td>
                          )}
                        </tr>
                      ))
                    ) : (
                      <tr>
                        <td colSpan={visibleColumnCount} className="px-4 py-12 text-center">
                          <div className="flex flex-col items-center gap-2">
                            <span className="text-slate-400 italic">No items found matching your filters.</span>
                            {(searchTerm || selectedPartNames.length > 0 || selectedDates.length > 0 || startDate || endDate) && (
                              <button 
                                onClick={resetFilters}
                                className="text-xs text-blue-600 hover:underline font-medium"
                              >
                                Clear all filters
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </React.Fragment>
        )}
      </React.Fragment>
    )}
  </div>
)}
        </div>
      </div>
    </main>

      {/* Floating Toggle Button for Zero Columns */}
      <button
        onClick={() => setHideZeroColumns(!hideZeroColumns)}
        className={cn(
          "fixed bottom-6 right-6 z-[200] flex items-center gap-2 px-4 py-2.5 rounded-full shadow-2xl transition-all text-xs font-bold border",
          hideZeroColumns 
            ? "bg-blue-600 text-white border-blue-500 hover:bg-blue-700" 
            : "bg-slate-800 text-white border-slate-700 hover:bg-slate-900"
        )}
        title={hideZeroColumns ? "Show all columns" : "Hide columns with zero transactions"}
      >
        {hideZeroColumns ? <Eye className="w-4 h-4" /> : <EyeOff className="w-4 h-4" />}
        <span>{hideZeroColumns ? "SHOW ALL COLUMNS" : "HIDE ZERO COLUMNS"}</span>
      </button>
    </div>
  );
}

function StatCard({ label, value, icon, color }: { label: string, value: string | number, icon: React.ReactNode, color: 'blue' | 'emerald' | 'rose' | 'amber' }) {
  const colors = {
    blue: "bg-blue-50 text-blue-600 border-blue-100",
    emerald: "bg-emerald-50 text-emerald-600 border-emerald-100",
    rose: "bg-rose-50 text-rose-600 border-rose-100",
    amber: "bg-amber-50 text-amber-600 border-amber-100"
  };

  return (
    <div className="bg-white border border-slate-200 rounded-xl p-5 shadow-sm">
      <div className="flex items-center justify-between mb-3">
        <div className={cn("p-2 rounded-lg border", colors[color])}>
          {icon}
        </div>
      </div>
      <div className="space-y-1">
        <p className="text-sm font-medium text-slate-500">{label}</p>
        <p className="text-2xl font-bold text-slate-900">{value}</p>
      </div>
    </div>
  );
}
