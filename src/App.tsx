import React, { useEffect, useState, useMemo, useRef } from 'react';
import Papa from 'papaparse';
import { domToPng } from 'modern-screenshot';
import * as XLSX from 'xlsx';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  LineChart,
  Line,
  Legend,
  LabelList,
} from 'recharts';
import {
  Users,
  Hospital,
  Calendar,
  CreditCard,
  TrendingUp,
  RefreshCw,
  Search,
  Filter,
  Download,
  AlertCircle,
  Plus,
  Save,
  ArrowLeft,
  Check,
  ChevronUp,
  ChevronDown,
  ArrowUpDown,
  FileSpreadsheet,
  FileText,
} from 'lucide-react';
import { format, parseISO, isValid, startOfDay } from 'date-fns';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// Utility for tailwind classes
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// Export helpers
const exportToPNG = async (elementId: string, fileName: string) => {
  const element = document.getElementById(elementId);
  if (!element) {
    console.error(`Element with id ${elementId} not found`);
    return;
  }
  
  try {
    // Small delay to ensure rendering is complete
    await new Promise(resolve => setTimeout(resolve, 100));
    
    const dataUrl = await domToPng(element, {
      backgroundColor: '#ffffff',
      scale: 2,
      quality: 1,
    });
    
    const link = document.createElement('a');
    link.href = dataUrl;
    link.download = `${fileName}.png`;
    link.click();
  } catch (err) {
    console.error('Error exporting to PNG:', err);
  }
};

const exportToExcel = (data: any[], fileName: string) => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};

const exportToCSV = (data: any[], fileName: string) => {
  const csv = Papa.unparse(data);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', `${fileName}.csv`);
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

// Helper to find column name regardless of case
const getCol = (row: PatientData, possibleNames: string[]) => {
  const keys = Object.keys(row);
  for (const name of possibleNames) {
    const found = keys.find(k => k.toUpperCase().includes(name.toUpperCase()));
    if (found) return row[found];
  }
  return '';
};

const getColByIndex = (row: PatientData, index: number) => {
  const keys = Object.keys(row);
  if (index >= 0 && index < keys.length) {
    return row[keys[index]];
  }
  return '';
};

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884d8', '#82ca9d', '#ffc658', '#8dd1e1'];

const ageGroupsList = [
  { id: '0-1h', label: '< 1 jam' },
  { id: '1-23h', label: '1-23 jam' },
  { id: '1-7d', label: '1-7 hr' },
  { id: '8-28d', label: '8-28 hr' },
  { id: '29d-3m', label: '29 hr - < 3 bln' },
  { id: '3-6m', label: '3 - < 6 bln' },
  { id: '6-11m', label: '6-11 bln' },
  { id: '1-4y', label: '1-4 th' },
  { id: '5-9y', label: '5-9 th' },
  { id: '10-14y', label: '10-14 th' },
  { id: '15-19y', label: '15-19 th' },
  { id: '20-24y', label: '20-24 th' },
  { id: '25-29y', label: '25-29 th' },
  { id: '30-34y', label: '30-34 th' },
  { id: '35-39y', label: '35-39 th' },
  { id: '40-44y', label: '40-44 th' },
  { id: '45-49y', label: '45-49 th' },
  { id: '50-54y', label: '50-54 th' },
  { id: '55-59y', label: '55-59 th' },
  { id: '60-64y', label: '60-64 th' },
  { id: '65-69y', label: '65-69 th' },
  { id: '70-74y', label: '70-74 th' },
  { id: '75-79y', label: '75-79 th' },
  { id: '80-84y', label: '80-84 th' },
  { id: '85y+', label: '>= 85 th' },
];

const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1cZNsBNBriMiZL_QnI5zt1HC-dpgPpNv9AbGGG5ZdLA4/export?format=csv&gid=0';

interface PatientData {
  [key: string]: string;
}

// Multi-select Combobox Component
const MultiSelectCombobox = ({ options, selected, onChange, label, placeholder }: { 
  options: string[], 
  selected: string[], 
  onChange: (selected: string[]) => void,
  label: string,
  placeholder: string
}) => {
  const [isOpen, setIsOpen] = useState(false);
  const [search, setSearch] = useState('');
  const dropdownRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const filteredOptions = options.filter(opt => 
    opt.toLowerCase().includes(search.toLowerCase())
  );

  const toggleOption = (option: string) => {
    if (selected.includes(option)) {
      onChange(selected.filter(item => item !== option));
    } else {
      onChange([...selected, option]);
    }
  };

  return (
    <div className="flex flex-col relative" ref={dropdownRef}>
      <label className="text-[10px] text-slate-400 font-bold uppercase mb-0.5 ml-1">{label}</label>
      <div 
        className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus-within:ring-2 focus-within:ring-blue-500 outline-none min-w-[150px] cursor-pointer flex items-center justify-between"
        onClick={() => setIsOpen(!isOpen)}
      >
        <span className="truncate max-w-[120px]">
          {selected.length === 0 ? placeholder : `${selected.length} dipilih`}
        </span>
        <ChevronDown className={`w-3 h-3 transition-transform ${isOpen ? 'rotate-180' : ''}`} />
      </div>

      {isOpen && (
        <div className="absolute top-full left-0 mt-1 w-full min-w-[200px] bg-white border border-slate-200 rounded-lg shadow-lg z-50 p-2">
          <input 
            type="text"
            className="w-full bg-slate-50 border border-slate-200 rounded-md px-2 py-1 text-xs mb-2 outline-none focus:ring-1 focus:ring-blue-500"
            placeholder="Cari..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            onClick={(e) => e.stopPropagation()}
          />
          <div className="max-h-[200px] overflow-y-auto space-y-1">
            {filteredOptions.length === 0 ? (
              <div className="text-[10px] text-slate-400 text-center py-2">Tidak ditemukan</div>
            ) : (
              filteredOptions.map(opt => (
                <div 
                  key={opt}
                  className={`flex items-center gap-2 px-2 py-1.5 rounded-md cursor-pointer hover:bg-slate-50 transition-colors ${selected.includes(opt) ? 'bg-blue-50 text-blue-700' : 'text-slate-600'}`}
                  onClick={(e) => {
                    e.stopPropagation();
                    toggleOption(opt);
                  }}
                >
                  <div className={`w-3 h-3 rounded border flex items-center justify-center ${selected.includes(opt) ? 'bg-blue-600 border-blue-600' : 'border-slate-300'}`}>
                    {selected.includes(opt) && <Check className="w-2 h-2 text-white" />}
                  </div>
                  <span className="text-[10px] font-medium">{opt}</span>
                </div>
              ))
            )}
          </div>
          {options.length > 0 && (
            <div className="mt-2 pt-2 border-t border-slate-100 flex flex-col gap-2">
              <div className="flex justify-between items-center">
                <button 
                  className="text-[10px] text-blue-600 font-bold hover:underline"
                  onClick={(e) => {
                    e.stopPropagation();
                    const newSelected = Array.from(new Set([...selected, ...filteredOptions]));
                    onChange(newSelected);
                  }}
                >
                  Pilih Semua
                </button>
                <button 
                  className="text-[10px] text-red-600 font-bold hover:underline"
                  onClick={(e) => {
                    e.stopPropagation();
                    if (search) {
                      // If searching, only remove the filtered options from selection
                      onChange(selected.filter(item => !filteredOptions.includes(item)));
                    } else {
                      // If not searching, clear everything
                      onChange([]);
                    }
                  }}
                >
                  Hapus Semua
                </button>
              </div>
              <button 
                className="w-full text-[10px] bg-slate-100 text-slate-600 font-bold py-1 rounded hover:bg-slate-200 transition-colors"
                onClick={(e) => {
                  e.stopPropagation();
                  setIsOpen(false);
                }}
              >
                Tutup
              </button>
            </div>
          )}
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [data, setData] = useState<PatientData[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<'overview' | 'diagnosa' | 'inputRL51' | 'rl52' | 'rl53' | 'periode' | 'programPrioritas'>('overview');
  const [selectedDiseaseForInput, setSelectedDiseaseForInput] = useState<any>(null);
  const [inputSearchTerm, setInputSearchTerm] = useState('');
  const [periodeSort, setPeriodeSort] = useState<{ key: string; direction: 'asc' | 'desc' }>({
    key: 'clinic',
    direction: 'asc'
  });
  const [periodeInsuranceSort, setPeriodeInsuranceSort] = useState<{ key: string; direction: 'asc' | 'desc' }>({
    key: 'insurance',
    direction: 'asc'
  });
  const [periodeAgeSort, setPeriodeAgeSort] = useState<{ key: string; direction: 'asc' | 'desc' }>({
    key: 'age',
    direction: 'asc'
  });
  const [periodePriorityProgramSort, setPeriodePriorityProgramSort] = useState<{ key: string; direction: 'asc' | 'desc' }>({
    key: 'program',
    direction: 'asc'
  });

  const handlePeriodeSort = (key: string) => {
    setPeriodeSort(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };
  const handlePeriodeInsuranceSort = (key: string) => {
    setPeriodeInsuranceSort(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };
  const handlePeriodeAgeSort = (key: string) => {
    setPeriodeAgeSort(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };
  const handlePeriodePriorityProgramSort = (key: string) => {
    setPeriodePriorityProgramSort(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedClinic, setSelectedClinic] = useState<string>('All');
  const [selectedYear, setSelectedYear] = useState<string>('All');
  const [selectedMonth, setSelectedMonth] = useState<string>('All');
  const [selectedGender, setSelectedGender] = useState<string>('All');
  const [selectedInsurance, setSelectedInsurance] = useState<string>('All');
  const [selectedCaseType, setSelectedCaseType] = useState<string>('All');
  const [selectedICDs, setSelectedICDs] = useState<string[]>([]);
  const [clinicLimit, setClinicLimit] = useState<'10' | '20' | 'All'>('10');
  const [showDataLabels, setShowDataLabels] = useState(false);
  const [inputYear, setInputYear] = useState('2026');
  const [inputMonth, setInputMonth] = useState('Januari');
  const [inputValues, setInputValues] = useState<Record<string, { L: number, P: number }>>({});
  const [totalVisitsInput, setTotalVisitsInput] = useState<{ L: number, P: number }>({ L: 0, P: 0 });

  const handleSelectDisease = (disease: any) => {
    setSelectedDiseaseForInput(disease);
    // Pre-fill with existing data from the spreadsheet stats
    const values: Record<string, { L: number, P: number }> = {};
    ageGroupsList.forEach(g => {
      values[g.id] = { 
        L: disease.ageGroups?.[g.id]?.L || 0, 
        P: disease.ageGroups?.[g.id]?.P || 0 
      };
    });
    setInputValues(values);
    setTotalVisitsInput({
      L: disease.totalVisits?.L || 0,
      P: disease.totalVisits?.P || 0
    });
  };

  const handleInputChange = (ageGroupId: string, gender: 'L' | 'P', value: string) => {
    const numValue = parseInt(value) || 0;
    setInputValues(prev => ({
      ...prev,
      [ageGroupId]: {
        ...prev[ageGroupId],
        [gender]: numValue
      }
    }));
  };

  const clinics = useMemo(() => {
    const set = new Set<string>();
    data.forEach(row => {
      const val = getCol(row, ['POLI', 'KLINIK', 'POLIKLINIK']);
      if (val) set.add(val);
    });
    return ['All', ...Array.from(set).sort()];
  }, [data]);

  const years = useMemo(() => {
    const set = new Set<string>();
    data.forEach(row => {
      const dateStr = getCol(row, ['TANGGAL', 'DATE', 'TGL']);
      if (dateStr) {
        const dateOnly = dateStr.split(' ')[0];
        const parts = dateOnly.split(/[-/.]/);
        if (parts.length === 3) {
          if (parts[0].length === 4) set.add(parts[0]);
          else if (parts[2].length === 4) set.add(parts[2]);
          else if (parts[2].length === 2) set.add('20' + parts[2]);
        }
      }
    });
    return ['All', ...Array.from(set).sort().reverse()];
  }, [data]);

  const insurances = useMemo(() => {
    const set = new Set<string>();
    data.forEach(row => {
      const val = getCol(row, ['PENJAMIN', 'ASURANSI', 'CARA BAYAR']);
      if (val) set.add(val);
    });
    return ['All', ...Array.from(set).sort()];
  }, [data]);

  const icds = useMemo(() => {
    const set = new Set<string>();
    data.forEach(row => {
      const val = getColByIndex(row, 12); // Column M
      if (val) set.add(val);
    });
    return Array.from(set).sort();
  }, [data]);

  const months = [
    'All', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];

  const fetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      Papa.parse(SHEET_URL, {
        download: true,
        header: true,
        skipEmptyLines: true,
        complete: (results) => {
          const rawData = results.data as PatientData[];
          
          // Check if we got HTML instead of CSV (common if not published to web)
          if (rawData.length > 0) {
            const firstVal = Object.values(rawData[0])[0];
            if (typeof firstVal === 'string' && (firstVal.includes('<!DOCTYPE') || firstVal.includes('<html'))) {
              setError('Data tidak valid. Pastikan Google Sheet telah "Dipublikasikan ke Web" sebagai CSV (File > Share > Publish to web).');
              setLoading(false);
              return;
            }
          }

          // Filter out rows that are completely empty
          const cleanedData = rawData.filter(row => 
            Object.values(row).some(val => val && val.trim() !== '')
          );
          
          if (cleanedData.length === 0) {
            setError('Data kosong. Pastikan spreadsheet memiliki data.');
          } else {
            setData(cleanedData);
          }
          setLoading(false);
        },
        error: (err) => {
          console.error('PapaParse Error:', err);
          setError('Gagal mengambil data dari Google Sheets. Pastikan link publik.');
          setLoading(false);
        },
      });
    } catch (err) {
      console.error('Fetch Error:', err);
      setError('Terjadi kesalahan saat mengambil data.');
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, []);

  const filteredData = useMemo(() => {
    return data.filter(row => {
      const clinic = getCol(row, ['POLI', 'KLINIK', 'POLIKLINIK']);
      const name = getCol(row, ['NAMA', 'PASIEN']);
      const insurance = getCol(row, ['PENJAMIN', 'ASURANSI', 'CARA BAYAR']);
      const gender = (getCol(row, ['JENIS KELAMIN', 'JK', 'SEX']) || '').toUpperCase().startsWith('L') ? 'L' : 'P';
      const dateStr = getCol(row, ['TANGGAL', 'DATE', 'TGL']);
      
      let rowYear = '';
      let rowMonth = '';
      if (dateStr) {
        try {
          // Handle various date formats
          // Try split by space first to remove time if present
          const dateOnly = dateStr.split(' ')[0];
          const parts = dateOnly.split(/[-/.]/);
          if (parts.length === 3) {
            if (parts[0].length === 4) { // YYYY-MM-DD
              rowYear = parts[0];
              rowMonth = parts[1];
            } else if (parts[2].length === 4) { // DD/MM/YYYY
              rowYear = parts[2];
              rowMonth = parts[1];
            } else if (parts[2].length === 2) { // DD/MM/YY - assume 20xx
              rowYear = '20' + parts[2];
              rowMonth = parts[1];
            }
          }
        } catch (e) {}
      }

      const matchesSearch = name?.toLowerCase().includes(searchTerm.toLowerCase()) || 
                           clinic?.toLowerCase().includes(searchTerm.toLowerCase());
      const matchesClinic = selectedClinic === 'All' || clinic === selectedClinic;
      const matchesYear = selectedYear === 'All' || rowYear === selectedYear;
      
      // Month mapping
      const monthMap: Record<string, string> = {
        'Januari': '01', 'Februari': '02', 'Maret': '03', 'April': '04',
        'Mei': '05', 'Juni': '06', 'Juli': '07', 'Agustus': '08',
        'September': '09', 'Oktober': '10', 'November': '11', 'Desember': '12'
      };
      const targetMonthCode = monthMap[selectedMonth];
      const matchesMonth = selectedMonth === 'All' || rowMonth === targetMonthCode;
      
      const matchesGender = selectedGender === 'All' || gender === selectedGender;
      const matchesInsurance = selectedInsurance === 'All' || insurance === selectedInsurance;
      
      const caseTypeVal = getColByIndex(row, 13); // Column N
      const caseType = caseTypeVal === 'v' ? 'Kasus Baru' : 'Kasus Lama';
      const matchesCaseType = selectedCaseType === 'All' || caseType === selectedCaseType;
      
      const icd = getColByIndex(row, 12);
      const matchesICD = selectedICDs.length === 0 || selectedICDs.includes(icd);

      return matchesSearch && matchesClinic && matchesYear && matchesMonth && matchesGender && matchesInsurance && matchesCaseType && matchesICD;
    });
  }, [data, searchTerm, selectedClinic, selectedYear, selectedMonth, selectedGender, selectedInsurance, selectedCaseType, selectedICDs]);

  // Sync Input RL 5.1 period with global filters when they change
  useEffect(() => {
    if (selectedYear !== 'All') setInputYear(selectedYear);
    if (selectedMonth !== 'All') setInputMonth(selectedMonth);
  }, [selectedYear, selectedMonth]);

  const getAgeGroupId = (row: PatientData) => {
    const ageCategory = (getCol(row, ['KATEGORI UMUR', 'KELOMPOK UMUR', 'AGE CATEGORY']) || '').toLowerCase();
    const normalized = ageCategory.trim();
    
    if (normalized.includes('0-28 hr') || normalized.includes('0-28 hari')) return '0-28d';
    if (normalized.includes('1-23 jam')) return '1-23h';
    if (normalized.includes('1-7 hr')) return '1-7d';
    if (normalized.includes('8-28 hr')) return '8-28d';
    if (normalized.includes('29 hr') || normalized.includes('< 3 bln')) return '29d-3m';
    if (normalized.includes('3 - < 6 bln')) return '3-6m';
    if (normalized.includes('6-11 bln')) return '6-11m';
    if (normalized.includes('1-4 th')) return '1-4y';
    if (normalized.includes('5-9 th')) return '5-9y';
    if (normalized.includes('10-14 th')) return '10-14y';
    if (normalized.includes('15-19 th')) return '15-19y';
    if (normalized.includes('20-24 th')) return '20-24y';
    if (normalized.includes('25-29 th')) return '25-29y';
    if (normalized.includes('30-34 th')) return '30-34y';
    if (normalized.includes('35-39 th')) return '35-39y';
    if (normalized.includes('40-44 th')) return '40-44y';
    if (normalized.includes('45-49 th')) return '45-49y';
    if (normalized.includes('50-54 th')) return '50-54y';
    if (normalized.includes('55-59 th')) return '55-59y';
    if (normalized.includes('60-64 th')) return '60-64y';
    if (normalized.includes('65-69 th')) return '65-69y';
    if (normalized.includes('70-74 th')) return '70-74y';
    if (normalized.includes('75-79 th')) return '75-79y';
    if (normalized.includes('80-84 th')) return '80-84y';
    if (normalized.includes('>= 85 th') || normalized.includes('85 th')) return '85y+';

    // Fallback to calculation
    const ageStr = getCol(row, ['UMUR', 'USIA', 'AGE']) || '0';
    const ageVal = parseInt(ageStr) || 0;
    const ageUnit = (getCol(row, ['SATUAN UMUR', 'UNIT']) || 'th').toLowerCase();
    
    if (ageUnit.includes('jam')) {
      if (ageVal < 1) return '0-1h';
      return '1-23h';
    } else if (ageUnit.includes('hr') || ageUnit.includes('hari')) {
      if (ageVal <= 7) return '1-7d';
      if (ageVal <= 28) return '8-28d';
      return '29d-3m';
    } else if (ageUnit.includes('bln') || ageUnit.includes('bulan')) {
      if (ageVal < 3) return '29d-3m';
      if (ageVal < 6) return '3-6m';
      return '6-11m';
    } else {
      if (ageVal < 1) return '6-11m';
      if (ageVal <= 4) return '1-4y';
      if (ageVal <= 9) return '5-9y';
      if (ageVal <= 14) return '10-14y';
      if (ageVal <= 19) return '15-19y';
      if (ageVal <= 24) return '20-24y';
      if (ageVal <= 29) return '25-29y';
      if (ageVal <= 34) return '30-34y';
      if (ageVal <= 39) return '35-39y';
      if (ageVal <= 44) return '40-44y';
      if (ageVal <= 49) return '45-49y';
      if (ageVal <= 54) return '50-54y';
      if (ageVal <= 59) return '55-59y';
      if (ageVal <= 64) return '60-64y';
      if (ageVal <= 69) return '65-69y';
      if (ageVal <= 74) return '70-74y';
      if (ageVal <= 79) return '75-79y';
      if (ageVal <= 84) return '80-84y';
      return '85y+';
    }
  };

  const diagnosisStats = useMemo(() => {
    if (filteredData.length === 0) return [];

    const ageGroups = [
      { id: '0-1h', label: '< 1 jam' },
      { id: '1-23h', label: '1-23 jam' },
      { id: '1-7d', label: '1-7 hr' },
      { id: '8-28d', label: '8-28 hr' },
      { id: '29d-3m', label: '29 hr - < 3 bln' },
      { id: '3-6m', label: '3 - < 6 bln' },
      { id: '6-11m', label: '6-11 bln' },
      { id: '1-4y', label: '1-4 th' },
      { id: '5-9y', label: '5-9 th' },
      { id: '10-14y', label: '10-14 th' },
      { id: '15-19y', label: '15-19 th' },
      { id: '20-24y', label: '20-24 th' },
      { id: '25-29y', label: '25-29 th' },
      { id: '30-34y', label: '30-34 th' },
      { id: '35-39y', label: '35-39 th' },
      { id: '40-44y', label: '40-44 th' },
      { id: '45-49y', label: '45-49 th' },
      { id: '50-54y', label: '50-54 th' },
      { id: '55-59y', label: '55-59 th' },
      { id: '60-64y', label: '60-64 th' },
      { id: '65-69y', label: '65-69 th' },
      { id: '70-74y', label: '70-74 th' },
      { id: '75-79y', label: '75-79 th' },
      { id: '80-84y', label: '80-84 th' },
      { id: '85y+', label: '>= 85 th' },
    ];

    const diagnosisMap: Record<string, any> = {};

    filteredData.forEach(row => {
      // Prioritize column names over indices for better robustness
      const icd = getCol(row, ['KODE ICD', 'ICD10', 'ICD 10', 'KODE']) || getColByIndex(row, 12) || '-';
      const diagnosis = getCol(row, ['DIAGNOSA', 'PENYAKIT', 'DIAGNOSIS', 'NAMA PENYAKIT']) || 'Unknown';
      const gender = (getCol(row, ['JENIS KELAMIN', 'JK', 'SEX']) || '').toUpperCase().startsWith('L') ? 'L' : 'P';
      
      // Check for 'v' in Kasus Baru column
      const isNewCaseVal = getCol(row, ['KASUS BARU', 'NEW CASE', 'BARU']) || getColByIndex(row, 13) || '';
      const isNewCase = isNewCaseVal.toLowerCase().trim() === 'v';
      
      // Check for age group column
      const ageGroupFromSheet = getCol(row, ['KELOMPOK UMUR', 'UMUR', 'AGE GROUP']) || getColByIndex(row, 19);
      
      let ageGroupId = '1-4y'; // Default
      
      if (ageGroupFromSheet) {
        // Map string from sheet to our IDs
        const normalized = ageGroupFromSheet.toLowerCase();
        if (normalized.includes('< 1 jam')) ageGroupId = '0-1h';
        else if (normalized.includes('1-23 jam')) ageGroupId = '1-23h';
        else if (normalized.includes('1-7 hr')) ageGroupId = '1-7d';
        else if (normalized.includes('8-28 hr')) ageGroupId = '8-28d';
        else if (normalized.includes('29 hr') || normalized.includes('< 3 bln')) ageGroupId = '29d-3m';
        else if (normalized.includes('3 - < 6 bln')) ageGroupId = '3-6m';
        else if (normalized.includes('6-11 bln')) ageGroupId = '6-11m';
        else if (normalized.includes('1-4 th')) ageGroupId = '1-4y';
        else if (normalized.includes('5-9 th')) ageGroupId = '5-9y';
        else if (normalized.includes('10-14 th')) ageGroupId = '10-14y';
        else if (normalized.includes('15-19 th')) ageGroupId = '15-19y';
        else if (normalized.includes('20-24 th')) ageGroupId = '20-24y';
        else if (normalized.includes('25-29 th')) ageGroupId = '25-29y';
        else if (normalized.includes('30-34 th')) ageGroupId = '30-34y';
        else if (normalized.includes('35-39 th')) ageGroupId = '35-39y';
        else if (normalized.includes('40-44 th')) ageGroupId = '40-44y';
        else if (normalized.includes('45-49 th')) ageGroupId = '45-49y';
        else if (normalized.includes('50-54 th')) ageGroupId = '50-54y';
        else if (normalized.includes('55-59 th')) ageGroupId = '55-59y';
        else if (normalized.includes('60-64 th')) ageGroupId = '60-64y';
        else if (normalized.includes('65-69 th')) ageGroupId = '65-69y';
        else if (normalized.includes('70-74 th')) ageGroupId = '70-74y';
        else if (normalized.includes('75-79 th')) ageGroupId = '75-79y';
        else if (normalized.includes('80-84 th')) ageGroupId = '80-84y';
        else if (normalized.includes('>= 85 th') || normalized.includes('85 th')) ageGroupId = '85y+';
        else {
          // Fallback to calculation if manual mapping fails
          const ageStr = getCol(row, ['UMUR', 'USIA', 'AGE']) || '0';
          const ageVal = parseInt(ageStr) || 0;
          const ageUnit = (getCol(row, ['SATUAN UMUR', 'UNIT']) || 'th').toLowerCase();
          
          if (ageUnit.includes('jam')) {
            if (ageVal < 1) ageGroupId = '0-1h';
            else ageGroupId = '1-23h';
          } else if (ageUnit.includes('hr') || ageUnit.includes('hari')) {
            if (ageVal <= 7) ageGroupId = '1-7d';
            else if (ageVal <= 28) ageGroupId = '8-28d';
            else ageGroupId = '29d-3m';
          } else if (ageUnit.includes('bln') || ageUnit.includes('bulan')) {
            if (ageVal < 3) ageGroupId = '29d-3m';
            else if (ageVal < 6) ageGroupId = '3-6m';
            else ageGroupId = '6-11m';
          } else {
            if (ageVal < 1) ageGroupId = '6-11m';
            else if (ageVal <= 4) ageGroupId = '1-4y';
            else if (ageVal <= 9) ageGroupId = '5-9y';
            else if (ageVal <= 14) ageGroupId = '10-14y';
            else if (ageVal <= 19) ageGroupId = '15-19y';
            else if (ageVal <= 24) ageGroupId = '20-24y';
            else if (ageVal <= 29) ageGroupId = '25-29y';
            else if (ageVal <= 34) ageGroupId = '30-34y';
            else if (ageVal <= 39) ageGroupId = '35-39y';
            else if (ageVal <= 44) ageGroupId = '40-44y';
            else if (ageVal <= 49) ageGroupId = '45-49y';
            else if (ageVal <= 54) ageGroupId = '50-54y';
            else if (ageVal <= 59) ageGroupId = '55-59y';
            else if (ageVal <= 64) ageGroupId = '60-64y';
            else if (ageVal <= 69) ageGroupId = '65-69y';
            else if (ageVal <= 74) ageGroupId = '70-74y';
            else if (ageVal <= 79) ageGroupId = '75-79y';
            else if (ageVal <= 84) ageGroupId = '80-84y';
            else ageGroupId = '85y+';
          }
        }
      }

      const key = `${icd}-${diagnosis}`;
      if (!diagnosisMap[key]) {
        diagnosisMap[key] = {
          icd,
          diagnosis,
          ageGroups: {},
          totalNew: { L: 0, P: 0, total: 0 },
          totalVisits: { L: 0, P: 0, total: 0 },
        };
        ageGroups.forEach(g => {
          diagnosisMap[key].ageGroups[g.id] = { L: 0, P: 0 };
        });
      }

      // Increment visits
      diagnosisMap[key].totalVisits[gender]++;
      diagnosisMap[key].totalVisits.total++;

      // Increment new cases if applicable
      if (isNewCase) {
        diagnosisMap[key].ageGroups[ageGroupId][gender]++;
        diagnosisMap[key].totalNew[gender]++;
        diagnosisMap[key].totalNew.total++;
      }
    });

    return Object.values(diagnosisMap).sort((a, b) => a.icd.localeCompare(b.icd));
  }, [filteredData]);

  const filteredDiagnosisData = useMemo(() => {
    return diagnosisStats.filter(item => 
      item.icd.toLowerCase().includes(searchTerm.toLowerCase()) || 
      item.diagnosis.toLowerCase().includes(searchTerm.toLowerCase())
    );
  }, [diagnosisStats, searchTerm]);

  const filteredInputDiseases = useMemo(() => {
    return diagnosisStats.filter(item => 
      item.icd.toLowerCase().includes(inputSearchTerm.toLowerCase()) || 
      item.diagnosis.toLowerCase().includes(inputSearchTerm.toLowerCase())
    );
  }, [diagnosisStats, inputSearchTerm]);

  const rl52Stats = useMemo(() => {
    if (filteredData.length === 0) return [];

    const diagnosisMap: Record<string, any> = {};

    filteredData.forEach(row => {
      // Use Column Z (index 25) and Column AA (index 26) as requested
      const icd = getColByIndex(row, 25) || '-';
      const diagnosis = getColByIndex(row, 26) || 'Unknown';
      const gender = (getCol(row, ['JENIS KELAMIN', 'JK', 'SEX']) || '').toUpperCase().startsWith('L') ? 'L' : 'P';
      
      const isNewCaseVal = getCol(row, ['KASUS BARU', 'NEW CASE', 'BARU']) || getColByIndex(row, 13) || '';
      const isNewCase = isNewCaseVal.toLowerCase().trim() === 'v';
      
      const key = `${icd}-${diagnosis}`;
      if (!diagnosisMap[key]) {
        diagnosisMap[key] = {
          icd,
          diagnosis,
          totalNew: { L: 0, P: 0, total: 0 },
          totalVisits: { L: 0, P: 0, total: 0 },
        };
      }

      // Increment visits
      diagnosisMap[key].totalVisits[gender]++;
      diagnosisMap[key].totalVisits.total++;

      // Increment new cases if applicable
      if (isNewCase) {
        diagnosisMap[key].totalNew[gender]++;
        diagnosisMap[key].totalNew.total++;
      }
    });

    return Object.values(diagnosisMap)
      .filter((item: any) => item.totalNew.total > 0 && item.icd !== '#N/A')
      .sort((a: any, b: any) => b.totalNew.total - a.totalNew.total)
      .slice(0, 10);
  }, [filteredData]);

  const rl53Stats = useMemo(() => {
    if (filteredData.length === 0) return [];

    const diagnosisMap: Record<string, any> = {};

    filteredData.forEach(row => {
      const icd = getColByIndex(row, 25) || '-';
      const diagnosis = getColByIndex(row, 26) || 'Unknown';
      const gender = (getCol(row, ['JENIS KELAMIN', 'JK', 'SEX']) || '').toUpperCase().startsWith('L') ? 'L' : 'P';
      
      const isNewCaseVal = getCol(row, ['KASUS BARU', 'NEW CASE', 'BARU']) || getColByIndex(row, 13) || '';
      const isNewCase = isNewCaseVal.toLowerCase().trim() === 'v';
      
      const key = `${icd}-${diagnosis}`;
      if (!diagnosisMap[key]) {
        diagnosisMap[key] = {
          icd,
          diagnosis,
          totalNew: { L: 0, P: 0, total: 0 },
          totalVisits: { L: 0, P: 0, total: 0 },
        };
      }

      diagnosisMap[key].totalVisits[gender]++;
      diagnosisMap[key].totalVisits.total++;

      if (isNewCase) {
        diagnosisMap[key].totalNew[gender]++;
        diagnosisMap[key].totalNew.total++;
      }
    });

    return Object.values(diagnosisMap)
      .filter((item: any) => item.totalVisits.total > 0 && item.icd !== '#N/A')
      .sort((a: any, b: any) => b.totalVisits.total - a.totalVisits.total)
      .slice(0, 10);
  }, [filteredData]);

  const periodeStats = useMemo(() => {
    if (filteredData.length === 0) return { 
      matrix: {}, clinics: [], 
      insuranceMatrix: {}, insuranceList: [],
      ageMatrix: {}, ageList: [],
      priorityProgramMatrix: {}, priorityProgramList: [],
      months: [], monthMap: {} 
    };

    const matrix: Record<string, Record<string, number>> = {};
    const insuranceMatrix: Record<string, Record<string, number>> = {};
    const ageMatrix: Record<string, Record<string, number>> = {};
    const priorityProgramMatrix: Record<string, Record<string, number>> = {};
    
    const clinicSet = new Set<string>();
    const insuranceSet = new Set<string>();
    const ageSet = new Set<string>();
    const priorityProgramSet = new Set<string>();

    const monthMap: Record<string, string> = {
      '01': 'Januari', '02': 'Februari', '03': 'Maret', '04': 'April',
      '05': 'Mei', '06': 'Juni', '07': 'Juli', '08': 'Agustus',
      '09': 'September', '10': 'Oktober', '11': 'November', '12': 'Desember'
    };
    const monthOrder = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'];

    filteredData.forEach(row => {
      const clinic = getCol(row, ['POLI', 'KLINIK', 'POLIKLINIK']) || 'Unknown';
      const insurance = getCol(row, ['PENJAMIN', 'ASURANSI', 'CARA BAYAR']) || 'Umum';
      const ageGroupId = getAgeGroupId(row);
      const priorityProgram = getColByIndex(row, 27) || 'Non Program';
      const dateStr = getCol(row, ['TANGGAL', 'DATE', 'TGL']);
      
      let rowMonth = '';
      if (dateStr) {
        const dateOnly = dateStr.split(' ')[0];
        const parts = dateOnly.split(/[-/.]/);
        if (parts.length === 3) {
          rowMonth = parts[1].padStart(2, '0');
        }
      }

      if (rowMonth && monthMap[rowMonth]) {
        // Clinic
        clinicSet.add(clinic);
        if (!matrix[clinic]) matrix[clinic] = {};
        matrix[clinic][rowMonth] = (matrix[clinic][rowMonth] || 0) + 1;

        // Insurance
        insuranceSet.add(insurance);
        if (!insuranceMatrix[insurance]) insuranceMatrix[insurance] = {};
        insuranceMatrix[insurance][rowMonth] = (insuranceMatrix[insurance][rowMonth] || 0) + 1;

        // Age
        ageSet.add(ageGroupId);
        if (!ageMatrix[ageGroupId]) ageMatrix[ageGroupId] = {};
        ageMatrix[ageGroupId][rowMonth] = (ageMatrix[ageGroupId][rowMonth] || 0) + 1;

        // Priority Program
        priorityProgramSet.add(priorityProgram);
        if (!priorityProgramMatrix[priorityProgram]) priorityProgramMatrix[priorityProgram] = {};
        priorityProgramMatrix[priorityProgram][rowMonth] = (priorityProgramMatrix[priorityProgram][rowMonth] || 0) + 1;
      }
    });

    const sortedClinics = Array.from(clinicSet).sort();
    const sortedInsurance = Array.from(insuranceSet).sort();
    const sortedPriorityPrograms = Array.from(priorityProgramSet).sort();
    
    // Custom sort for age groups
    const ageOrder = [
      '0-1h', '1-23h', '1-7d', '8-28d', '29d-3m', '3-6m', '6-11m',
      '1-4y', '5-9y', '10-14y', '15-19y', '20-24y', '25-29y', '30-34y',
      '35-39y', '40-44y', '45-49y', '50-54y', '55-59y', '60-64y',
      '65-69y', '70-74y', '75-79y', '80-84y', '85y+'
    ];
    const sortedAgeGroups = Array.from(ageSet).sort((a, b) => {
      return ageOrder.indexOf(a) - ageOrder.indexOf(b);
    });

    // Only show months that have data
    const activeMonths = monthOrder.filter(m => 
      sortedClinics.some(c => (matrix[c]?.[m] || 0) > 0) ||
      sortedInsurance.some(i => (insuranceMatrix[i]?.[m] || 0) > 0) ||
      sortedAgeGroups.some(a => (ageMatrix[a]?.[m] || 0) > 0) ||
      sortedPriorityPrograms.some(p => (priorityProgramMatrix[p]?.[m] || 0) > 0)
    );
    
    const displayMonths = activeMonths.length > 0 ? activeMonths : monthOrder;

    return { 
      matrix, clinics: sortedClinics, 
      insuranceMatrix, insuranceList: sortedInsurance,
      ageMatrix, ageList: sortedAgeGroups,
      priorityProgramMatrix, priorityProgramList: sortedPriorityPrograms,
      months: displayMonths, monthMap 
    };
  }, [filteredData]);

  const sortedPeriodeClinics = useMemo(() => {
    const { clinics, matrix } = periodeStats;
    if (!clinics.length) return [];

    const sorted = [...clinics];
    const { key, direction } = periodeSort;

    sorted.sort((a, b) => {
      let valA: any;
      let valB: any;

      if (key === 'clinic') {
        valA = a;
        valB = b;
      } else if (key === 'total') {
        valA = Object.values(matrix[a] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
        valB = Object.values(matrix[b] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
      } else {
        valA = matrix[a]?.[key] || 0;
        valB = matrix[b]?.[key] || 0;
      }

      if (typeof valA === 'string' && typeof valB === 'string') {
        return direction === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
      }

      if (valA < valB) return direction === 'asc' ? -1 : 1;
      if (valA > valB) return direction === 'asc' ? 1 : -1;
      return 0;
    });

    return sorted;
  }, [periodeStats, periodeSort]);

  const sortedPeriodeInsurance = useMemo(() => {
    const { insuranceList, insuranceMatrix } = periodeStats;
    if (!insuranceList.length) return [];

    const sorted = [...insuranceList];
    const { key, direction } = periodeInsuranceSort;

    sorted.sort((a, b) => {
      let valA: any;
      let valB: any;

      if (key === 'insurance') {
        valA = a;
        valB = b;
      } else if (key === 'total') {
        valA = Object.values(insuranceMatrix[a] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
        valB = Object.values(insuranceMatrix[b] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
      } else {
        valA = insuranceMatrix[a]?.[key] || 0;
        valB = insuranceMatrix[b]?.[key] || 0;
      }

      if (typeof valA === 'string' && typeof valB === 'string') {
        return direction === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
      }

      if (valA < valB) return direction === 'asc' ? -1 : 1;
      if (valA > valB) return direction === 'asc' ? 1 : -1;
      return 0;
    });

    return sorted;
  }, [periodeStats, periodeInsuranceSort]);

  const sortedPeriodeAge = useMemo(() => {
    const { ageList, ageMatrix } = periodeStats;
    if (!ageList.length) return [];

    const sorted = [...ageList];
    const { key, direction } = periodeAgeSort;

    sorted.sort((a, b) => {
      let valA: any;
      let valB: any;

      if (key === 'age') {
        valA = a;
        valB = b;
      } else if (key === 'total') {
        valA = Object.values(ageMatrix[a] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
        valB = Object.values(ageMatrix[b] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
      } else {
        valA = ageMatrix[a]?.[key] || 0;
        valB = ageMatrix[b]?.[key] || 0;
      }

      if (typeof valA === 'string' && typeof valB === 'string') {
        return direction === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
      }

      if (valA < valB) return direction === 'asc' ? -1 : 1;
      if (valA > valB) return direction === 'asc' ? 1 : -1;
      return 0;
    });

    return sorted;
  }, [periodeStats, periodeAgeSort]);

  const sortedPeriodePriorityProgram = useMemo(() => {
    const { priorityProgramList, priorityProgramMatrix } = periodeStats;
    if (!priorityProgramList.length) return [];

    const sorted = [...priorityProgramList];
    const { key, direction } = periodePriorityProgramSort;

    sorted.sort((a, b) => {
      let valA: any;
      let valB: any;

      if (key === 'program') {
        valA = a;
        valB = b;
      } else if (key === 'total') {
        valA = Object.values(priorityProgramMatrix[a] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
        valB = Object.values(priorityProgramMatrix[b] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
      } else {
        valA = priorityProgramMatrix[a]?.[key] || 0;
        valB = priorityProgramMatrix[b]?.[key] || 0;
      }

      if (typeof valA === 'string' && typeof valB === 'string') {
        return direction === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
      }

      if (valA < valB) return direction === 'asc' ? -1 : 1;
      if (valA > valB) return direction === 'asc' ? 1 : -1;
      return 0;
    });

    return sorted;
  }, [periodeStats, periodePriorityProgramSort]);

  // Export handlers for tables
  const handleExportPeriodeClinic = (format: 'excel' | 'csv') => {
    const { clinics, matrix, months, monthMap } = periodeStats;
    const exportData = clinics.map(clinic => {
      const row: any = { 'Nama Poliklinik': clinic };
      let total = 0;
      months.forEach(m => {
        const val = matrix[clinic]?.[m] || 0;
        row[monthMap[m]] = val;
        total += val;
      });
      row['Total'] = total;
      return row;
    });
    if (format === 'excel') exportToExcel(exportData, 'kunjungan-poliklinik-periode');
    else exportToCSV(exportData, 'kunjungan-poliklinik-periode');
  };

  const handleExportPeriodeInsurance = (format: 'excel' | 'csv') => {
    const { insuranceList, insuranceMatrix, months, monthMap } = periodeStats;
    const exportData = insuranceList.map(insurance => {
      const row: any = { 'Nama Penjamin': insurance };
      let total = 0;
      months.forEach(m => {
        const val = insuranceMatrix[insurance]?.[m] || 0;
        row[monthMap[m]] = val;
        total += val;
      });
      row['Total'] = total;
      return row;
    });
    if (format === 'excel') exportToExcel(exportData, 'kunjungan-penjamin-periode');
    else exportToCSV(exportData, 'kunjungan-penjamin-periode');
  };

  const handleExportPeriodeAge = (format: 'excel' | 'csv') => {
    const { ageList, ageMatrix, months, monthMap } = periodeStats;
    const exportData = ageList.map(ageGroup => {
      const label = ageGroupsList.find(g => g.id === ageGroup)?.label || ageGroup;
      const row: any = { 'Kategori Umur': label };
      let total = 0;
      months.forEach(m => {
        const val = ageMatrix[ageGroup]?.[m] || 0;
        row[monthMap[m]] = val;
        total += val;
      });
      row['Total'] = total;
      return row;
    });
    if (format === 'excel') exportToExcel(exportData, 'kunjungan-umur-periode');
    else exportToCSV(exportData, 'kunjungan-umur-periode');
  };

  const handleExportPeriodePriorityProgram = (format: 'excel' | 'csv') => {
    const { priorityProgramList, priorityProgramMatrix, months, monthMap } = periodeStats;
    const exportData = priorityProgramList.map(program => {
      const row: any = { 'Nama Program Prioritas': program };
      let total = 0;
      months.forEach(m => {
        const val = priorityProgramMatrix[program]?.[m] || 0;
        row[monthMap[m]] = val;
        total += val;
      });
      row['Total'] = total;
      return row;
    });
    if (format === 'excel') exportToExcel(exportData, 'data-program-prioritas-periode');
    else exportToCSV(exportData, 'data-program-prioritas-periode');
  };

  const handleExportRL51 = (format: 'excel' | 'csv') => {
    const exportData = filteredDiagnosisData.map(row => {
      const exportRow: any = {
        'Kode ICD': row.icd,
        'Diagnosis': row.diagnosis
      };
      
      ageGroupsList.forEach(group => {
        exportRow[`${group.label} L`] = row.ageGroups[group.id].L;
        exportRow[`${group.label} P`] = row.ageGroups[group.id].P;
      });
      
      exportRow['Total Baru L'] = row.totalNew.L;
      exportRow['Total Baru P'] = row.totalNew.P;
      exportRow['Total Baru'] = row.totalNew.total;
      exportRow['Total Kunjungan L'] = row.totalVisits.L;
      exportRow['Total Kunjungan P'] = row.totalVisits.P;
      exportRow['Total Kunjungan'] = row.totalVisits.total;
      
      return exportRow;
    });
    
    if (format === 'excel') exportToExcel(exportData, 'rl51-morbiditas');
    else exportToCSV(exportData, 'rl51-morbiditas');
  };

  const handleExportRL52 = (format: 'excel' | 'csv') => {
    const exportData = rl52Stats.map((row: any) => ({
      'Kode ICD-10': row.icd,
      'Diagnosis Penyakit': row.diagnosis,
      'Kasus Baru L': row.totalNew.L,
      'Kasus Baru P': row.totalNew.P,
      'Total Kasus Baru': row.totalNew.total,
      'Kunjungan L': row.totalVisits.L,
      'Kunjungan P': row.totalVisits.P,
      'Total Kunjungan': row.totalVisits.total
    }));
    if (format === 'excel') exportToExcel(exportData, 'rl52-10-besar-kasus-baru');
    else exportToCSV(exportData, 'rl52-10-besar-kasus-baru');
  };

  const handleExportRL53 = (format: 'excel' | 'csv') => {
    const exportData = rl53Stats.map((row: any) => ({
      'Kode ICD-10': row.icd,
      'Diagnosis Penyakit': row.diagnosis,
      'Kasus Baru L': row.totalNew.L,
      'Kasus Baru P': row.totalNew.P,
      'Total Kasus Baru': row.totalNew.total,
      'Kunjungan L': row.totalVisits.L,
      'Kunjungan P': row.totalVisits.P,
      'Total Kunjungan': row.totalVisits.total
    }));
    if (format === 'excel') exportToExcel(exportData, 'rl53-10-besar-kunjungan');
    else exportToCSV(exportData, 'rl53-10-besar-kunjungan');
  };

  const periodeChartData = useMemo(() => {
    const { clinics, matrix, months, monthMap } = periodeStats;
    if (!months.length) return [];

    // Get top 5 clinics by total visits in this period
    const topClinics = [...clinics]
      .sort((a, b) => {
        const totalA = Object.values(matrix[a] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
        const totalB = Object.values(matrix[b] || {}).reduce((sum: number, v: any) => sum + (v as number), 0);
        return totalB - totalA;
      })
      .slice(0, 5);

    return months.map(m => {
      const row: any = { name: monthMap[m] };
      topClinics.forEach(clinic => {
        row[clinic] = matrix[clinic]?.[m] || 0;
      });
      return row;
    });
  }, [periodeStats]);

  const stats = useMemo(() => {
    if (filteredData.length === 0) return null;

    const clinicCounts: Record<string, number> = {};
    const insuranceCounts: Record<string, number> = {};
    const dateCounts: Record<string, number> = {};
    const monthCounts: Record<string, number> = {};
    const genderCounts: Record<string, number> = { 'L': 0, 'P': 0 };
    const ageCategoryCounts: Record<string, number> = {};
    const dayCounts: Record<string, number> = {};
    const priorityProgramCounts: Record<string, number> = {};

    const MONTH_NAMES = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];

    const DAY_NAMES = [
      'Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'
    ];

    filteredData.forEach(row => {
      const clinic = getCol(row, ['POLI', 'KLINIK', 'POLIKLINIK']) || 'Unknown';
      const insurance = getCol(row, ['PENJAMIN', 'ASURANSI', 'CARA BAYAR']) || 'Umum';
      const gender = (getCol(row, ['JENIS KELAMIN', 'JK', 'SEX']) || '').toUpperCase().startsWith('L') ? 'L' : 'P';
      const dateStr = getCol(row, ['TANGGAL', 'DATE', 'TGL']);
      
      // Age Category
      const ageGroupFromSheet = getCol(row, ['KELOMPOK UMUR', 'UMUR', 'AGE GROUP']) || getColByIndex(row, 19) || 'Lainnya';
      ageCategoryCounts[ageGroupFromSheet] = (ageCategoryCounts[ageGroupFromSheet] || 0) + 1;

      const priorityProgram = getColByIndex(row, 27) || 'Non Program';
      priorityProgramCounts[priorityProgram] = (priorityProgramCounts[priorityProgram] || 0) + 1;

      clinicCounts[clinic] = (clinicCounts[clinic] || 0) + 1;
      insuranceCounts[insurance] = (insuranceCounts[insurance] || 0) + 1;
      genderCounts[gender]++;

      if (dateStr) {
        try {
          const dateOnly = dateStr.split(' ')[0];
          dateCounts[dateOnly] = (dateCounts[dateOnly] || 0) + 1;
          
          const parts = dateOnly.split(/[-/.]/);
          if (parts.length === 3) {
            let year = '';
            let month = '';
            let day = '';
            
            if (parts[0].length === 4) { // YYYY-MM-DD
              year = parts[0];
              month = parts[1];
              day = parts[2];
            } else if (parts[2].length === 4 || parts[2].length === 2) { // DD/MM/YYYY
              year = parts[2].length === 2 ? '20' + parts[2] : parts[2];
              month = parts[1];
              day = parts[0];
            }

            if (year && month) {
              const monthIdx = parseInt(month) - 1;
              const monthName = `${MONTH_NAMES[monthIdx]} ${year}`;
              monthCounts[monthName] = (monthCounts[monthName] || 0) + 1;
              
              const dateObj = new Date(parseInt(year), monthIdx, parseInt(day));
              if (!isNaN(dateObj.getTime())) {
                const dayName = DAY_NAMES[dateObj.getDay()];
                dayCounts[dayName] = (dayCounts[dayName] || 0) + 1;
              }
            }
          }
        } catch (e) {}
      }
    });

    const patientsByClinic = Object.entries(clinicCounts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const patientsByInsurance = Object.entries(insuranceCounts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    const patientsByDate = Object.entries(dateCounts)
      .map(([date, count]) => ({ date, count }))
      .sort((a, b) => a.date.localeCompare(b.date))
      .slice(-15);

    const patientsByMonth = Object.entries(monthCounts)
      .map(([name, value]) => ({ name, value }));
      // Sort by month/year logic could be better but for now let's use insertion order or simple sort
      // Actually, let's sort by year-month if possible

    const patientsByGender = [
      { name: 'Laki-laki', value: genderCounts['L'], color: '#2563eb' },
      { name: 'Perempuan', value: genderCounts['P'], color: '#db2777' }
    ];

    const patientsByAgeCategory = Object.entries(ageCategoryCounts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 10);

    const DAY_ORDER = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
    const patientsByDay = DAY_ORDER.map(day => ({
      name: day,
      value: dayCounts[day] || 0
    }));

    const patientsByPriorityProgram = Object.entries(priorityProgramCounts)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    return {
      totalPatients: filteredData.length,
      patientsByClinic,
      patientsByInsurance,
      patientsByDate,
      patientsByMonth,
      patientsByGender,
      patientsByAgeCategory,
      patientsByDay,
      patientsByPriorityProgram,
      clinics: ['All', ...Object.keys(clinicCounts).sort()],
    };
  }, [filteredData]);

  if (loading) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-4">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mb-4"></div>
        <p className="text-slate-600 font-medium animate-pulse">Memuat Data Poliklinik...</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-4">
        <div className="bg-white p-8 rounded-2xl shadow-xl max-w-md w-full text-center border border-red-100">
          <AlertCircle className="w-16 h-16 text-red-500 mx-auto mb-4" />
          <h2 className="text-2xl font-bold text-slate-800 mb-2">Oops! Ada Masalah</h2>
          <p className="text-slate-600 mb-6">{error}</p>
          <button 
            onClick={fetchData}
            className="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-3 rounded-xl transition-all flex items-center justify-center gap-2"
          >
            <RefreshCw className="w-5 h-5" /> Coba Lagi
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 font-sans pb-12">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="px-4 sm:px-6 lg:px-8 h-20 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2 rounded-xl">
              <Hospital className="w-6 h-6 text-white" />
            </div>
            <div>
              <h1 className="text-xl font-bold tracking-tight">Dashboard Poliklinik</h1>
              <p className="text-xs text-slate-500 font-medium uppercase tracking-wider">Monitoring Rawat Jalan 2026</p>
            </div>
          </div>
          
          <div className="hidden md:flex items-center gap-4">
            <nav className="flex items-center bg-slate-100 p-1 rounded-xl mr-4">
              <button 
                onClick={() => setActiveTab('overview')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'overview' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                Overview
              </button>
              <button 
                onClick={() => setActiveTab('diagnosa')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'diagnosa' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                Diagnosa
              </button>
              <button 
                onClick={() => setActiveTab('inputRL51')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'inputRL51' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                Input RL 5.1
              </button>
              <button 
                onClick={() => setActiveTab('rl52')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'rl52' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                RL 5.2
              </button>
              <button 
                onClick={() => setActiveTab('rl53')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'rl53' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                RL 5.3
              </button>
              <button 
                onClick={() => setActiveTab('periode')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'periode' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                Periode
              </button>
              <button 
                onClick={() => setActiveTab('programPrioritas')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'programPrioritas' ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                )}
              >
                Program Prioritas
              </button>
            </nav>
            <div className="relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
              <input 
                type="text" 
                placeholder="Cari pasien atau poli..." 
                className="pl-10 pr-4 py-2 bg-slate-100 border-transparent focus:bg-white focus:ring-2 focus:ring-blue-500 rounded-lg text-sm w-64 transition-all"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
            <button 
              onClick={fetchData}
              className="p-2 text-slate-500 hover:bg-slate-100 rounded-lg transition-colors"
              title="Refresh Data"
            >
              <RefreshCw className={cn("w-5 h-5", loading && "animate-spin")} />
            </button>
          </div>
        </div>

      {/* Global Filters Bar */}
      <div className="bg-white border-b border-slate-200 py-3 sticky top-20 z-10 shadow-sm">
        <div className="px-4 sm:px-6 lg:px-8">
          <div className="flex flex-wrap items-center gap-4 text-sm">
            <div className="flex items-center gap-2 text-blue-600 font-bold bg-blue-50 px-3 py-1.5 rounded-lg border border-blue-100">
              <Filter className="w-4 h-4" />
              <span className="text-[10px] uppercase tracking-wider">Global Filters</span>
            </div>
            
            {/* Clinic Filter */}
            <div className="flex flex-col">
              <label className="text-[10px] text-slate-400 font-bold uppercase mb-0.5 ml-1">Poliklinik</label>
              <select 
                className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none min-w-[120px]"
                value={selectedClinic}
                onChange={(e) => setSelectedClinic(e.target.value)}
              >
                {clinics.map(c => <option key={c} value={c}>{c === 'All' ? 'Semua Poli' : c}</option>)}
              </select>
            </div>

            {/* Year Filter */}
            <div className="flex flex-col">
              <label className="text-[10px] text-slate-400 font-bold uppercase mb-0.5 ml-1">Tahun</label>
              <select 
                className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none min-w-[80px]"
                value={selectedYear}
                onChange={(e) => setSelectedYear(e.target.value)}
              >
                {years.map(y => <option key={y} value={y}>{y === 'All' ? 'Semua Tahun' : y}</option>)}
              </select>
            </div>

            {/* Month Filter */}
            <div className="flex flex-col">
              <label className="text-[10px] text-slate-400 font-bold uppercase mb-0.5 ml-1">Bulan</label>
              <select 
                className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none min-w-[100px]"
                value={selectedMonth}
                onChange={(e) => setSelectedMonth(e.target.value)}
              >
                {months.map(m => <option key={m} value={m}>{m === 'All' ? 'Semua Bulan' : m}</option>)}
              </select>
            </div>

            {/* ICD Filter */}
            <MultiSelectCombobox 
              options={icds}
              selected={selectedICDs}
              onChange={setSelectedICDs}
              label="ICD (Kolom M)"
              placeholder="Pilih ICD"
            />

            {/* Show Data Labels Checkbox */}
            <div className="flex items-center gap-2 ml-auto bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200 self-end">
              <input 
                type="checkbox" 
                id="showLabels"
                className="w-4 h-4 text-blue-600 rounded border-slate-300 focus:ring-blue-500 cursor-pointer"
                checked={showDataLabels}
                onChange={(e) => setShowDataLabels(e.target.checked)}
              />
              <label htmlFor="showLabels" className="text-xs font-bold text-slate-600 cursor-pointer select-none">Tampilkan Angka</label>
            </div>

            {/* Gender Filter */}
            <div className="flex flex-col">
              <label className="text-[10px] text-slate-400 font-bold uppercase mb-0.5 ml-1">Jenis Kelamin</label>
              <select 
                className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none min-w-[100px]"
                value={selectedGender}
                onChange={(e) => setSelectedGender(e.target.value)}
              >
                <option value="All">Semua JK</option>
                <option value="L">Laki-laki</option>
                <option value="P">Perempuan</option>
              </select>
            </div>

            {/* Insurance Filter */}
            <div className="flex flex-col">
              <label className="text-[10px] text-slate-400 font-bold uppercase mb-0.5 ml-1">Penjamin</label>
              <select 
                className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none min-w-[120px]"
                value={selectedInsurance}
                onChange={(e) => setSelectedInsurance(e.target.value)}
              >
                {insurances.map(i => <option key={i} value={i}>{i === 'All' ? 'Semua Penjamin' : i}</option>)}
              </select>
            </div>

            {/* Case Type Filter */}
            <div className="flex flex-col">
              <label className="text-[10px] text-slate-400 font-bold uppercase mb-0.5 ml-1">Jenis Kasus</label>
              <select 
                className="bg-slate-50 border border-slate-200 rounded-lg px-3 py-1.5 text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none min-w-[120px]"
                value={selectedCaseType}
                onChange={(e) => setSelectedCaseType(e.target.value)}
              >
                <option value="All">Semua Kasus</option>
                <option value="Kasus Baru">Kasus Baru</option>
                <option value="Kasus Lama">Kasus Lama</option>
              </select>
            </div>

            <button 
              onClick={() => {
                setSelectedClinic('All');
                setSelectedYear('2026');
                setSelectedMonth('All');
                setSelectedGender('All');
                setSelectedInsurance('All');
                setSelectedCaseType('All');
                setSelectedICDs([]);
                setSearchTerm('');
              }}
              className="mt-auto mb-0.5 px-3 py-1.5 text-xs font-bold text-red-600 hover:bg-red-50 rounded-lg transition-colors flex items-center gap-1"
            >
              <RefreshCw className="w-3 h-3" /> Reset
            </button>
          </div>
        </div>
      </div>
      </header>

      <main className="px-4 sm:px-6 lg:px-8 py-8">
        {/* Active Filters Summary */}
        <div className="mb-6 flex flex-wrap items-center gap-2">
          <span className="text-xs font-bold text-slate-400 uppercase tracking-widest mr-2">Menampilkan Data:</span>
          {selectedClinic !== 'All' && (
            <span className="bg-blue-100 text-blue-700 px-3 py-1 rounded-full text-[10px] font-bold border border-blue-200">Poli: {selectedClinic}</span>
          )}
          {selectedYear !== 'All' && (
            <span className="bg-emerald-100 text-emerald-700 px-3 py-1 rounded-full text-[10px] font-bold border border-emerald-200">Tahun: {selectedYear}</span>
          )}
          {selectedMonth !== 'All' && (
            <span className="bg-amber-100 text-amber-700 px-3 py-1 rounded-full text-[10px] font-bold border border-amber-200">Bulan: {selectedMonth}</span>
          )}
          {selectedGender !== 'All' && (
            <span className="bg-purple-100 text-purple-700 px-3 py-1 rounded-full text-[10px] font-bold border border-purple-200">JK: {selectedGender === 'L' ? 'Laki-laki' : 'Perempuan'}</span>
          )}
          {selectedInsurance !== 'All' && (
            <span className="bg-pink-100 text-pink-700 px-3 py-1 rounded-full text-[10px] font-bold border border-pink-200">Penjamin: {selectedInsurance}</span>
          )}
          {selectedCaseType !== 'All' && (
            <span className="bg-orange-100 text-orange-700 px-3 py-1 rounded-full text-[10px] font-bold border border-orange-200">Kasus: {selectedCaseType}</span>
          )}
          {selectedICDs.length > 0 && (
            <span className="bg-indigo-100 text-indigo-700 px-3 py-1 rounded-full text-[10px] font-bold border border-indigo-200">ICD: {selectedICDs.length} dipilih</span>
          )}
          {searchTerm && (
            <span className="bg-slate-100 text-slate-700 px-3 py-1 rounded-full text-[10px] font-bold border border-slate-200">Cari: "{searchTerm}"</span>
          )}
          {selectedClinic === 'All' && selectedYear === 'All' && selectedMonth === 'All' && selectedGender === 'All' && selectedInsurance === 'All' && selectedCaseType === 'All' && selectedICDs.length === 0 && !searchTerm && (
            <span className="text-xs text-slate-400 italic">Semua data (Tanpa filter)</span>
          )}
        </div>

        {activeTab === 'overview' ? (
          <>
            {/* Filters & Stats Summary */}
            <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-8">
              <div className="md:col-span-3 grid grid-cols-1 sm:grid-cols-3 gap-4">
                {/* Stat Cards */}
                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center gap-4">
                  <div className="bg-blue-50 p-3 rounded-xl">
                    <Users className="w-6 h-6 text-blue-600" />
                  </div>
                  <div>
                    <p className="text-sm text-slate-500 font-medium">Total Pasien</p>
                    <p className="text-2xl font-bold">{stats?.totalPatients.toLocaleString()}</p>
                  </div>
                </div>

                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center gap-4">
                  <div className="bg-emerald-50 p-3 rounded-xl">
                    <TrendingUp className="w-6 h-6 text-emerald-600" />
                  </div>
                  <div>
                    <p className="text-sm text-slate-500 font-medium">Poli Teraktif</p>
                    <p className="text-lg font-bold truncate max-w-[150px]">{stats?.patientsByClinic[0]?.name || '-'}</p>
                  </div>
                </div>

                <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center gap-4">
                  <div className="bg-amber-50 p-3 rounded-xl">
                    <CreditCard className="w-6 h-6 text-amber-600" />
                  </div>
                  <div>
                    <p className="text-sm text-slate-500 font-medium">Penjamin Utama</p>
                    <p className="text-lg font-bold truncate max-w-[150px]">{stats?.patientsByInsurance[0]?.name || '-'}</p>
                  </div>
                </div>
              </div>

              {/* Clinic Filter */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                <div className="flex items-center gap-2 mb-3 text-slate-500">
                  <Filter className="w-4 h-4" />
                  <span className="text-xs font-bold uppercase tracking-wider">Filter Poliklinik</span>
                </div>
                <select 
                  className="w-full bg-slate-50 border-slate-200 rounded-lg text-sm focus:ring-blue-500 focus:border-blue-500"
                  value={selectedClinic}
                  onChange={(e) => setSelectedClinic(e.target.value)}
                >
                  {stats?.clinics.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>
            </div>

            {/* Charts Grid */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8 mb-8">
              {/* 1. Tren Kunjungan Bulanan */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200" id="chart-trend">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <TrendingUp className="w-5 h-5 text-blue-600" />
                    Tren Kunjungan Bulanan
                  </h3>
                  <button 
                    onClick={() => exportToPNG('chart-trend', 'tren-kunjungan')}
                    className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                    title="Download PNG"
                  >
                    <Download className="w-4 h-4" />
                  </button>
                </div>
                <div className="h-[300px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <LineChart data={stats?.patientsByMonth}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis 
                        dataKey="name" 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fontSize: 10, fill: '#94a3b8' }} 
                        dy={10} 
                      />
                      <YAxis 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fontSize: 10, fill: '#94a3b8' }} 
                      />
                      <Tooltip 
                        contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} 
                      />
                      <Line 
                        type="monotone" 
                        dataKey="value" 
                        stroke="#2563eb" 
                        strokeWidth={3} 
                        dot={{ r: 4, fill: '#2563eb', strokeWidth: 2, stroke: '#fff' }} 
                        activeDot={{ r: 6, strokeWidth: 0 }}
                      >
                        {showDataLabels && <LabelList dataKey="value" position="top" offset={10} style={{ fontSize: '10px', fontWeight: 'bold', fill: '#2563eb' }} />}
                      </Line>
                    </LineChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* 2. Kunjungan Jenis Kelamin */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200" id="chart-gender">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <Users className="w-5 h-5 text-pink-600" />
                    Kunjungan Jenis Kelamin
                  </h3>
                  <button 
                    onClick={() => exportToPNG('chart-gender', 'kunjungan-gender')}
                    className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-pink-600"
                    title="Download PNG"
                  >
                    <Download className="w-4 h-4" />
                  </button>
                </div>
                <div className="h-[300px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie
                        data={stats?.patientsByGender}
                        cx="50%" 
                        cy="50%" 
                        innerRadius={60} 
                        outerRadius={80} 
                        paddingAngle={5} 
                        dataKey="value"
                        label={showDataLabels ? ({ name, value }) => `${name}: ${value}` : false}
                      >
                        {stats?.patientsByGender.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={entry.color} />
                        ))}
                      </Pie>
                      <Tooltip 
                        contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} 
                      />
                      <Legend 
                        verticalAlign="bottom" 
                        height={36} 
                        iconType="circle" 
                        wrapperStyle={{ fontSize: '12px', fontWeight: 'bold' }} 
                      />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* 3. Kunjungan Berdasarkan Kategori Umur */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200" id="chart-age">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <Users className="w-5 h-5 text-emerald-600" />
                    Kunjungan Berdasarkan Kategori Umur
                  </h3>
                  <div className="flex items-center gap-2">
                    <span className="text-xs text-slate-400 font-medium">Top 10 Kategori</span>
                    <button 
                      onClick={() => exportToPNG('chart-age', 'kunjungan-umur')}
                      className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-emerald-600"
                      title="Download PNG"
                    >
                      <Download className="w-4 h-4" />
                    </button>
                  </div>
                </div>
                <div className="h-[300px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={stats?.patientsByAgeCategory} layout="vertical">
                      <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#f1f5f9" />
                      <XAxis type="number" hide />
                      <YAxis 
                        dataKey="name" 
                        type="category" 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fontSize: 10, fill: '#64748b' }} 
                        width={120} 
                      />
                      <Tooltip 
                        cursor={{ fill: '#f8fafc' }} 
                        contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} 
                      />
                      <Bar dataKey="value" fill="#10b981" radius={[0, 4, 4, 0]}>
                        {showDataLabels && <LabelList dataKey="value" position="right" offset={5} style={{ fontSize: '10px', fontWeight: 'bold', fill: '#10b981' }} />}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* 4. Kunjungan Berdasarkan Hari Kunjungan */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200" id="chart-day">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <Calendar className="w-5 h-5 text-amber-600" />
                    Kunjungan Berdasarkan Hari Kunjungan
                  </h3>
                  <button 
                    onClick={() => exportToPNG('chart-day', 'kunjungan-hari')}
                    className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-amber-600"
                    title="Download PNG"
                  >
                    <Download className="w-4 h-4" />
                  </button>
                </div>
                <div className="h-[300px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={stats?.patientsByDay}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis 
                        dataKey="name" 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fontSize: 10, fill: '#64748b' }} 
                      />
                      <YAxis 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fontSize: 10, fill: '#64748b' }} 
                      />
                      <Tooltip 
                        cursor={{ fill: '#f8fafc' }} 
                        contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} 
                      />
                      <Bar dataKey="value" fill="#f59e0b" radius={[4, 4, 0, 0]}>
                        {showDataLabels && <LabelList dataKey="value" position="top" offset={5} style={{ fontSize: '10px', fontWeight: 'bold', fill: '#f59e0b' }} />}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* 5. Grafik Poliklinik (Top 10, 20, Semua) */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 lg:col-span-2" id="chart-clinic">
                <div className="flex flex-col md:flex-row md:items-center justify-between mb-6 gap-4">
                  <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <Hospital className="w-5 h-5 text-blue-600" />
                    Kunjungan Berdasarkan Poliklinik
                  </h3>
                  <div className="flex items-center gap-4">
                    <div className="flex items-center bg-slate-100 p-1 rounded-xl">
                      {(['10', '20', 'All'] as const).map((limit) => (
                        <button
                          key={limit}
                          onClick={() => setClinicLimit(limit)}
                          className={cn(
                            "px-4 py-1 rounded-lg text-xs font-bold transition-all",
                            clinicLimit === limit ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"
                          )}
                        >
                          {limit === 'All' ? 'Semua' : `Top ${limit}`}
                        </button>
                      ))}
                    </div>
                    <button 
                      onClick={() => exportToPNG('chart-clinic', 'kunjungan-poliklinik')}
                      className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                      title="Download PNG"
                    >
                      <Download className="w-4 h-4" />
                    </button>
                  </div>
                </div>
                <div className="h-[400px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart 
                      data={clinicLimit === 'All' ? stats?.patientsByClinic : stats?.patientsByClinic.slice(0, parseInt(clinicLimit))} 
                      margin={{ top: 5, right: 30, left: 20, bottom: 60 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                      <XAxis 
                        dataKey="name" 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fontSize: 10, fill: '#64748b' }}
                        angle={-45}
                        textAnchor="end"
                        interval={0}
                      />
                      <YAxis 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fontSize: 10, fill: '#64748b' }} 
                      />
                      <Tooltip 
                        cursor={{ fill: '#f8fafc' }} 
                        contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }} 
                      />
                      <Bar dataKey="value" fill="#3b82f6" radius={[4, 4, 0, 0]}>
                        {showDataLabels && <LabelList dataKey="value" position="top" offset={5} style={{ fontSize: '10px', fontWeight: 'bold', fill: '#3b82f6' }} />}
                        {stats?.patientsByClinic.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-8">
              {/* Recent Patients Table */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 lg:col-span-2 overflow-hidden" id="table-recent">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <Users className="w-5 h-5 text-slate-600" />
                    Data Pasien Terkini
                  </h3>
                  <div className="flex items-center gap-4">
                    <span className="text-xs font-medium text-slate-400">{filteredData.length} baris ditemukan</span>
                    <div className="flex items-center gap-2">
                      <button 
                        onClick={() => exportToPNG('table-recent', 'pasien-terkini')}
                        className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                        title="Download PNG"
                      >
                        <Download className="w-4 h-4" />
                      </button>
                      <button 
                        onClick={() => exportToExcel(filteredData, 'pasien-terkini')}
                        className="flex items-center gap-1 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-colors"
                        title="Export Excel"
                      >
                        <Download className="w-3 h-3" />
                        Excel
                      </button>
                      <button 
                        onClick={() => exportToCSV(filteredData, 'pasien-terkini')}
                        className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-bold hover:bg-slate-200 transition-colors"
                        title="Export CSV"
                      >
                        <Download className="w-3 h-3" />
                        CSV
                      </button>
                    </div>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm">
                    <thead>
                      <tr className="border-b border-slate-100">
                        <th className="pb-4 font-semibold text-slate-500">Tanggal</th>
                        <th className="pb-4 font-semibold text-slate-500">Nama Pasien</th>
                        <th className="pb-4 font-semibold text-slate-500">Poliklinik</th>
                        <th className="pb-4 font-semibold text-slate-500">Penjamin</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {filteredData.length === 0 ? (
                        <tr>
                          <td colSpan={4} className="py-12 text-center text-slate-400 italic">
                            Tidak ada data pasien yang ditemukan untuk filter ini.
                          </td>
                        </tr>
                      ) : (
                        filteredData.slice(0, 10).map((row, i) => (
                          <tr key={i} className="hover:bg-slate-50 transition-colors">
                            <td className="py-4 text-slate-600 font-mono text-xs">
                              {getCol(row, ['TANGGAL', 'DATE', 'TGL']) || '-'}
                            </td>
                            <td className="py-4 font-medium text-slate-800">
                              {getCol(row, ['NAMA', 'PASIEN']) || 'N/A'}
                            </td>
                            <td className="py-4">
                              <span className="inline-flex items-center px-2 py-1 rounded-md bg-blue-50 text-blue-700 text-[10px] font-bold uppercase tracking-wider">
                                {getCol(row, ['POLI', 'KLINIK', 'POLIKLINIK']) || '-'}
                              </span>
                            </td>
                            <td className="py-4 text-slate-500">
                              {getCol(row, ['PENJAMIN', 'ASURANSI', 'CARA BAYAR']) || '-'}
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                </div>
                {filteredData.length > 10 && (
                  <div className="mt-4 pt-4 border-top border-slate-50 text-center">
                    <p className="text-xs text-slate-400">Menampilkan 10 dari {filteredData.length} pasien</p>
                  </div>
                )}
              </div>
            </div>
          </>
        ) : activeTab === 'diagnosa' ? (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden" id="table-rl51">
            <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
              <div>
                <h3 className="text-lg font-bold text-slate-800">Laporan Morbiditas Pasien Rawat Jalan</h3>
                <p className="text-xs text-slate-500">Jumlah Kasus Baru Menurut Kelompok Umur & Jenis Kelamin</p>
              </div>
              <div className="flex items-center gap-4">
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => exportToPNG('table-rl51', 'rl51-morbiditas')}
                    className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                    title="Download PNG"
                  >
                    <Download className="w-4 h-4" />
                  </button>
                  <button 
                    onClick={() => handleExportRL51('excel')}
                    className="flex items-center gap-1 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-colors"
                  >
                    <Download className="w-3 h-3" /> Excel
                  </button>
                  <button 
                    onClick={() => handleExportRL51('csv')}
                    className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-bold hover:bg-slate-200 transition-colors"
                  >
                    <Download className="w-3 h-3" /> CSV
                  </button>
                </div>
                <div className="text-[10px] font-bold text-slate-400 uppercase tracking-widest flex gap-4">
                  <span>* L = Laki-laki</span>
                  <span>* P = Perempuan</span>
                </div>
              </div>
            </div>
            
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-[10px]">
                <thead>
                  <tr className="bg-slate-100 border-b border-slate-200">
                    <th rowSpan={3} className="p-2 border-r border-slate-200 min-w-[60px] text-center">Kode ICD</th>
                    <th rowSpan={3} className="p-2 border-r border-slate-200 min-w-[200px]">Diagnosis Penyakit</th>
                    <th colSpan={ageGroupsList.length * 2} className="p-2 border-b border-slate-200 text-center uppercase tracking-wider font-bold">
                      Jumlah Kasus Baru Menurut Kelompok Umur & Jenis Kelamin
                    </th>
                    <th colSpan={3} className="p-2 border-b border-slate-200 border-l border-slate-200 text-center font-bold">
                      Jumlah Kasus Baru
                    </th>
                    <th colSpan={3} className="p-2 border-b border-slate-200 border-l border-slate-200 text-center font-bold">
                      Jumlah Kunjungan
                    </th>
                  </tr>
                  <tr className="bg-slate-50 border-b border-slate-200">
                    {ageGroupsList.map(group => (
                      <th key={group.id} colSpan={2} className="p-1 border-r border-slate-200 text-center whitespace-nowrap min-w-[60px]">
                        {group.label}
                      </th>
                    ))}
                    <th colSpan={3} className="p-1 border-r border-slate-200 border-l border-slate-200 text-center">Menurut JK</th>
                    <th colSpan={3} className="p-1 border-l border-slate-200 text-center">Total</th>
                  </tr>
                  <tr className="bg-white border-b border-slate-200 font-bold">
                    {ageGroupsList.map(group => (
                      <React.Fragment key={group.id}>
                        <th className="p-1 border-r border-slate-100 text-center w-8">L</th>
                        <th className="p-1 border-r border-slate-200 text-center w-8">P</th>
                      </React.Fragment>
                    ))}
                    <th className="p-1 border-r border-slate-100 text-center w-8 bg-slate-50">L</th>
                    <th className="p-1 border-r border-slate-100 text-center w-8 bg-slate-50">P</th>
                    <th className="p-1 border-r border-slate-200 text-center w-10 bg-slate-100">Total</th>
                    <th className="p-1 border-r border-slate-100 text-center w-8 bg-slate-50">L</th>
                    <th className="p-1 border-r border-slate-100 text-center w-8 bg-slate-50">P</th>
                    <th className="p-1 text-center w-10 bg-slate-100">Total</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredDiagnosisData.length === 0 ? (
                    <tr>
                      <td colSpan={ageGroupsList.length * 2 + 8} className="p-8 text-center text-slate-400 italic">
                        Tidak ada data diagnosa yang ditemukan.
                      </td>
                    </tr>
                  ) : (
                    filteredDiagnosisData.map((row, idx) => (
                      <tr key={idx} className="hover:bg-blue-50/30 transition-colors">
                        <td className="p-2 border-r border-slate-100 font-mono text-center font-medium">{row.icd}</td>
                        <td className="p-2 border-r border-slate-100 font-medium text-slate-800">{row.diagnosis}</td>
                        {ageGroupsList.map(group => (
                          <React.Fragment key={group.id}>
                            <td className="p-1 border-r border-slate-100 text-center text-slate-600">
                              {row.ageGroups[group.id].L || ''}
                            </td>
                            <td className="p-1 border-r border-slate-100 text-center text-slate-600">
                              {row.ageGroups[group.id].P || ''}
                            </td>
                          </React.Fragment>
                        ))}
                        <td className="p-1 border-r border-slate-100 text-center font-bold bg-slate-50/50">{row.totalNew.L || ''}</td>
                        <td className="p-1 border-r border-slate-100 text-center font-bold bg-slate-50/50">{row.totalNew.P || ''}</td>
                        <td className="p-1 border-r border-slate-100 text-center font-black bg-slate-100/50">{row.totalNew.total || ''}</td>
                        <td className="p-1 border-r border-slate-100 text-center font-bold bg-slate-50/50">{row.totalVisits.L || ''}</td>
                        <td className="p-1 border-r border-slate-100 text-center font-bold bg-slate-50/50">{row.totalVisits.P || ''}</td>
                        <td className="p-1 text-center font-black bg-slate-100/50">{row.totalVisits.total || ''}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
            
            <div className="p-4 bg-slate-50 border-t border-slate-100 text-[10px] text-slate-400 italic">
              ** hr = hari, bln = bulan, th = tahun
            </div>
          </div>
        ) : activeTab === 'rl52' ? (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden" id="table-rl52">
            <div className="p-6 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <h3 className="text-lg font-bold text-slate-800">RL 5.2 10 Besar Kasus Baru Penyakit Rawat Jalan</h3>
              <div className="flex items-center gap-2">
                <button 
                  onClick={() => exportToPNG('table-rl52', 'rl52-10-besar-kasus-baru')}
                  className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                  title="Download PNG"
                >
                  <Download className="w-4 h-4" />
                </button>
                <button 
                  onClick={() => handleExportRL52('excel')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-colors"
                >
                  <Download className="w-3 h-3" /> Excel
                </button>
                <button 
                  onClick={() => handleExportRL52('csv')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-bold hover:bg-slate-200 transition-colors"
                >
                  <Download className="w-3 h-3" /> CSV
                </button>
              </div>
            </div>
            
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr className="bg-[#608282] text-white">
                    <th rowSpan={2} className="p-3 border border-slate-200 text-center w-12">No.</th>
                    <th rowSpan={2} className="p-3 border border-slate-200 text-center w-24">Kode ICD-10</th>
                    <th rowSpan={2} className="p-3 border border-slate-200 text-center">Diagnosis Penyakit</th>
                    <th colSpan={3} className="p-2 border border-slate-200 text-center font-bold">
                      Jumlah Kasus Baru Menurut Jenis Kelamin
                    </th>
                    <th colSpan={3} className="p-2 border border-slate-200 text-center font-bold">
                      Jumlah Kunjungan
                    </th>
                  </tr>
                  <tr className="bg-[#608282] text-white">
                    <th className="p-2 border border-slate-200 text-center w-24">Laki-Laki</th>
                    <th className="p-2 border border-slate-200 text-center w-24">Perempuan</th>
                    <th className="p-2 border border-slate-200 text-center w-24 bg-[#507272]">Total</th>
                    <th className="p-2 border border-slate-200 text-center w-24">Laki-Laki</th>
                    <th className="p-2 border border-slate-200 text-center w-24">Perempuan</th>
                    <th className="p-2 border border-slate-200 text-center w-24 bg-[#507272]">Total</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {rl52Stats.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="p-8 text-center text-slate-400 italic">
                        Tidak ada data untuk laporan RL 5.2.
                      </td>
                    </tr>
                  ) : (
                    rl52Stats.map((row: any, idx: number) => (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        <td className="p-3 border border-slate-100 text-center">{idx + 1}</td>
                        <td className="p-3 border border-slate-100 text-center font-mono">{row.icd}</td>
                        <td className="p-3 border border-slate-100 font-medium text-slate-800">{row.diagnosis}</td>
                        <td className="p-3 border border-slate-100 text-center">{row.totalNew.L || 0}</td>
                        <td className="p-3 border border-slate-100 text-center">{row.totalNew.P || 0}</td>
                        <td className="p-3 border border-slate-100 text-center font-bold bg-slate-50">{row.totalNew.total || 0}</td>
                        <td className="p-3 border border-slate-100 text-center">{row.totalVisits.L || 0}</td>
                        <td className="p-3 border border-slate-100 text-center">{row.totalVisits.P || 0}</td>
                        <td className="p-3 border border-slate-100 text-center font-bold bg-slate-50">{row.totalVisits.total || 0}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        ) : activeTab === 'rl53' ? (
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden" id="table-rl53">
            <div className="p-6 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <h3 className="text-lg font-bold text-slate-800">RL 5.3 10 Besar Kunjungan Penyakit Rawat Jalan</h3>
              <div className="flex items-center gap-2">
                <button 
                  onClick={() => exportToPNG('table-rl53', 'rl53-10-besar-kunjungan')}
                  className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                  title="Download PNG"
                >
                  <Download className="w-4 h-4" />
                </button>
                <button 
                  onClick={() => handleExportRL53('excel')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-colors"
                >
                  <Download className="w-3 h-3" /> Excel
                </button>
                <button 
                  onClick={() => handleExportRL53('csv')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-bold hover:bg-slate-200 transition-colors"
                >
                  <Download className="w-3 h-3" /> CSV
                </button>
              </div>
            </div>
            
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr className="bg-[#608282] text-white">
                    <th rowSpan={2} className="p-3 border border-slate-200 text-center w-12">No.</th>
                    <th rowSpan={2} className="p-3 border border-slate-200 text-center w-24">Kode ICD-10</th>
                    <th rowSpan={2} className="p-3 border border-slate-200 text-center">Diagnosis Penyakit</th>
                    <th colSpan={3} className="p-2 border border-slate-200 text-center font-bold">
                      Jumlah Kasus Baru Menurut Jenis Kelamin
                    </th>
                    <th colSpan={3} className="p-2 border border-slate-200 text-center font-bold">
                      Jumlah Kunjungan
                    </th>
                  </tr>
                  <tr className="bg-[#608282] text-white">
                    <th className="p-2 border border-slate-200 text-center w-24">Laki-Laki</th>
                    <th className="p-2 border border-slate-200 text-center w-24">Perempuan</th>
                    <th className="p-2 border border-slate-200 text-center w-24 bg-[#507272]">Total</th>
                    <th className="p-2 border border-slate-200 text-center w-24">Laki-Laki</th>
                    <th className="p-2 border border-slate-200 text-center w-24">Perempuan</th>
                    <th className="p-2 border border-slate-200 text-center w-24 bg-[#507272]">Total</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {rl53Stats.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="p-8 text-center text-slate-400 italic">
                        Tidak ada data untuk laporan RL 5.3.
                      </td>
                    </tr>
                  ) : (
                    rl53Stats.map((row: any, idx: number) => (
                      <tr key={idx} className="hover:bg-slate-50 transition-colors">
                        <td className="p-3 border border-slate-100 text-center">{idx + 1}</td>
                        <td className="p-3 border border-slate-100 text-center font-mono">{row.icd}</td>
                        <td className="p-3 border border-slate-100 font-medium text-slate-800">{row.diagnosis}</td>
                        <td className="p-3 border border-slate-100 text-center">{row.totalNew.L || 0}</td>
                        <td className="p-3 border border-slate-100 text-center">{row.totalNew.P || 0}</td>
                        <td className="p-3 border border-slate-100 text-center font-bold bg-slate-50">{row.totalNew.total || 0}</td>
                        <td className="p-3 border border-slate-100 text-center">{row.totalVisits.L || 0}</td>
                        <td className="p-3 border border-slate-100 text-center">{row.totalVisits.P || 0}</td>
                        <td className="p-3 border border-slate-100 text-center font-bold bg-slate-50">{row.totalVisits.total || 0}</td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        ) : activeTab === 'periode' ? (
          <div className="space-y-8">
            {/* Trend Chart for Periode */}
            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200" id="chart-periode-trend">
              <div className="flex items-center justify-between mb-6">
                <div>
                  <h3 className="text-lg font-bold text-slate-800">Tren Kunjungan 5 Poliklinik Terbanyak</h3>
                  <p className="text-xs text-slate-500">Perbandingan kunjungan bulanan antar poliklinik</p>
                </div>
                <button 
                  onClick={() => exportToPNG('chart-periode-trend', 'tren-kunjungan-periode')}
                  className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                  title="Download PNG"
                >
                  <Download className="w-4 h-4" />
                </button>
              </div>
              <div className="h-[400px] w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <LineChart data={periodeChartData}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                    <XAxis 
                      dataKey="name" 
                      axisLine={false} 
                      tickLine={false} 
                      tick={{ fill: '#64748b', fontSize: 12 }}
                      dy={10}
                    />
                    <YAxis 
                      axisLine={false} 
                      tickLine={false} 
                      tick={{ fill: '#64748b', fontSize: 12 }}
                    />
                    <Tooltip 
                      contentStyle={{ 
                        borderRadius: '12px', 
                        border: 'none', 
                        boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)',
                        padding: '12px'
                      }} 
                    />
                    <Legend />
                    {Object.keys(periodeChartData[0] || {}).filter(k => k !== 'name').map((key, index) => (
                      <Line 
                        key={key}
                        type="monotone" 
                        dataKey={key} 
                        stroke={COLORS[index % COLORS.length]} 
                        strokeWidth={3}
                        dot={{ r: 4, strokeWidth: 2, fill: '#fff' }}
                        activeDot={{ r: 6, strokeWidth: 0 }}
                      />
                    ))}
                  </LineChart>
                </ResponsiveContainer>
              </div>
            </div>

            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden" id="table-periode-clinic">
              <div className="p-6 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <div>
                <h3 className="text-lg font-bold text-slate-800">Jumlah Kunjungan Poliklinik Per Periode</h3>
                <p className="text-xs text-slate-500">Rekapitulasi kunjungan bulanan per poliklinik</p>
              </div>
              <div className="flex items-center gap-2">
                <button 
                  onClick={() => exportToPNG('table-periode-clinic', 'kunjungan-poliklinik-periode')}
                  className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                  title="Download PNG"
                >
                  <Download className="w-4 h-4" />
                </button>
                <button 
                  onClick={() => handleExportPeriodeClinic('excel')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-colors"
                >
                  <Download className="w-3 h-3" /> Excel
                </button>
                <button 
                  onClick={() => handleExportPeriodeClinic('csv')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-bold hover:bg-slate-200 transition-colors"
                >
                  <Download className="w-3 h-3" /> CSV
                </button>
              </div>
            </div>
            
            <div className="overflow-auto max-h-[600px]">
              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-800 text-white">
                    <th className="p-3 border border-slate-700 text-center sticky left-0 top-0 bg-slate-800 z-40 w-[40px]">No</th>
                    <th 
                      className="p-3 border border-slate-700 text-left sticky left-[40px] top-0 bg-slate-800 z-30 min-w-[180px] cursor-pointer hover:bg-slate-700 transition-colors"
                      onClick={() => handlePeriodeSort('clinic')}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span>Nama Poliklinik</span>
                        {periodeSort.key === 'clinic' ? (
                          periodeSort.direction === 'asc' ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />
                        ) : (
                          <ArrowUpDown className="w-3 h-3 opacity-30" />
                        )}
                      </div>
                    </th>
                    {periodeStats.months.map(m => (
                      <th 
                        key={m} 
                        className="p-3 border border-slate-700 text-center min-w-[90px] cursor-pointer hover:bg-slate-700 transition-colors sticky top-0 bg-slate-800 z-20"
                        onClick={() => handlePeriodeSort(m)}
                      >
                        <div className="flex flex-col items-center gap-1">
                          <span>{periodeStats.monthMap[m]}</span>
                          {periodeSort.key === m ? (
                            periodeSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                          ) : (
                            <ArrowUpDown className="w-2 h-2 opacity-30" />
                          )}
                        </div>
                      </th>
                    ))}
                    <th 
                      className="p-3 border border-slate-700 text-center bg-slate-700 font-bold min-w-[100px] cursor-pointer hover:bg-slate-600 transition-colors sticky top-0 z-20"
                      onClick={() => handlePeriodeSort('total')}
                    >
                      <div className="flex flex-col items-center gap-1">
                        <span>Total</span>
                        {periodeSort.key === 'total' ? (
                          periodeSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                        ) : (
                          <ArrowUpDown className="w-2 h-2 opacity-30" />
                        )}
                      </div>
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortedPeriodeClinics.length === 0 ? (
                    <tr>
                      <td colSpan={periodeStats.months.length + 3} className="p-8 text-center text-slate-400 italic">
                        Tidak ada data kunjungan untuk periode ini.
                      </td>
                    </tr>
                  ) : (
                    <>
                      {sortedPeriodeClinics.map((clinic, idx) => {
                        let rowTotal = 0;
                        return (
                          <tr key={idx} className="hover:bg-blue-50/50 transition-colors even:bg-slate-50/30">
                            <td className="p-3 border border-slate-100 text-center font-medium text-slate-400 sticky left-0 bg-inherit z-10">
                              {idx + 1}
                            </td>
                            <td className="p-3 border border-slate-100 font-bold text-slate-700 sticky left-[40px] bg-inherit z-10 shadow-[2px_0_5px_rgba(0,0,0,0.03)]">
                              {clinic}
                            </td>
                            {periodeStats.months.map(m => {
                              const val = periodeStats.matrix[clinic]?.[m] || 0;
                              rowTotal += val;
                              return (
                                <td key={m} className={cn(
                                  "p-3 border border-slate-100 text-center",
                                  val > 0 ? "text-blue-600 font-semibold" : "text-slate-300"
                                )}>
                                  {val || '-'}
                                </td>
                              );
                            })}
                            <td className="p-3 border border-slate-100 text-center font-black bg-slate-100/50 text-slate-900">
                              {rowTotal.toLocaleString()}
                            </td>
                          </tr>
                        );
                      })}
                      {/* Column Totals Row */}
                      <tr className="bg-slate-800 font-black text-white border-t-2 border-slate-700">
                        <td colSpan={2} className="p-3 border border-slate-700 sticky left-0 bg-slate-800 z-10 text-right pr-6">
                          TOTAL KESELURUHAN
                        </td>
                        {periodeStats.months.map(m => {
                          let colTotal = 0;
                          periodeStats.clinics.forEach(c => {
                            colTotal += (periodeStats.matrix[c]?.[m] || 0);
                          });
                          return (
                            <td key={m} className="p-3 border border-slate-700 text-center">
                              {colTotal.toLocaleString()}
                            </td>
                          );
                        })}
                        <td className="p-3 border border-slate-700 text-center bg-slate-900 text-blue-400 text-sm">
                          {(() => {
                            let grandTotal = 0;
                            periodeStats.clinics.forEach(c => {
                              periodeStats.months.forEach(m => {
                                grandTotal += (periodeStats.matrix[c]?.[m] || 0);
                              });
                            });
                            return grandTotal.toLocaleString();
                          })()}
                        </td>
                      </tr>
                    </>
                  )}
                </tbody>
              </table>
            </div>
          </div>

            {/* Insurance Table */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden mt-8" id="table-periode-insurance">
              <div className="p-6 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <div>
                <h3 className="text-lg font-bold text-slate-800">Jumlah Kunjungan Berdasarkan Penjamin Per Periode</h3>
                <p className="text-xs text-slate-500">Rekapitulasi kunjungan bulanan per penjamin/asuransi</p>
              </div>
              <div className="flex items-center gap-2">
                <button 
                  onClick={() => exportToPNG('table-periode-insurance', 'kunjungan-penjamin-periode')}
                  className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                  title="Download PNG"
                >
                  <Download className="w-4 h-4" />
                </button>
                <button 
                  onClick={() => handleExportPeriodeInsurance('excel')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-colors"
                >
                  <Download className="w-3 h-3" /> Excel
                </button>
                <button 
                  onClick={() => handleExportPeriodeInsurance('csv')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-bold hover:bg-slate-200 transition-colors"
                >
                  <Download className="w-3 h-3" /> CSV
                </button>
              </div>
            </div>
            
            <div className="overflow-auto max-h-[600px]">
              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-800 text-white">
                    <th className="p-3 border border-slate-700 text-center sticky left-0 top-0 bg-slate-800 z-40 w-[40px]">No</th>
                    <th 
                      className="p-3 border border-slate-700 text-left sticky left-[40px] top-0 bg-slate-800 z-30 min-w-[180px] cursor-pointer hover:bg-slate-700 transition-colors"
                      onClick={() => handlePeriodeInsuranceSort('insurance')}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span>Nama Penjamin</span>
                        {periodeInsuranceSort.key === 'insurance' ? (
                          periodeInsuranceSort.direction === 'asc' ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />
                        ) : (
                          <ArrowUpDown className="w-3 h-3 opacity-30" />
                        )}
                      </div>
                    </th>
                    {periodeStats.months.map(m => (
                      <th 
                        key={m} 
                        className="p-3 border border-slate-700 text-center min-w-[90px] cursor-pointer hover:bg-slate-700 transition-colors sticky top-0 bg-slate-800 z-20"
                        onClick={() => handlePeriodeInsuranceSort(m)}
                      >
                        <div className="flex flex-col items-center gap-1">
                          <span>{periodeStats.monthMap[m]}</span>
                          {periodeInsuranceSort.key === m ? (
                            periodeInsuranceSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                          ) : (
                            <ArrowUpDown className="w-2 h-2 opacity-30" />
                          )}
                        </div>
                      </th>
                    ))}
                    <th 
                      className="p-3 border border-slate-700 text-center bg-slate-700 font-bold min-w-[100px] cursor-pointer hover:bg-slate-600 transition-colors sticky top-0 z-20"
                      onClick={() => handlePeriodeInsuranceSort('total')}
                    >
                      <div className="flex flex-col items-center gap-1">
                        <span>Total</span>
                        {periodeInsuranceSort.key === 'total' ? (
                          periodeInsuranceSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                        ) : (
                          <ArrowUpDown className="w-2 h-2 opacity-30" />
                        )}
                      </div>
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortedPeriodeInsurance.length === 0 ? (
                    <tr>
                      <td colSpan={periodeStats.months.length + 3} className="p-8 text-center text-slate-400 italic">
                        Tidak ada data kunjungan untuk periode ini.
                      </td>
                    </tr>
                  ) : (
                    <>
                      {sortedPeriodeInsurance.map((insurance, idx) => {
                        let rowTotal = 0;
                        return (
                          <tr key={idx} className="hover:bg-green-50/50 transition-colors even:bg-slate-50/30">
                            <td className="p-3 border border-slate-100 text-center font-medium text-slate-400 sticky left-0 bg-inherit z-10">
                              {idx + 1}
                            </td>
                            <td className="p-3 border border-slate-100 font-bold text-slate-700 sticky left-[40px] bg-inherit z-10 shadow-[2px_0_5px_rgba(0,0,0,0.03)]">
                              {insurance}
                            </td>
                            {periodeStats.months.map(m => {
                              const val = periodeStats.insuranceMatrix[insurance]?.[m] || 0;
                              rowTotal += val;
                              return (
                                <td key={m} className={cn(
                                  "p-3 border border-slate-100 text-center",
                                  val > 0 ? "text-green-600 font-semibold" : "text-slate-300"
                                )}>
                                  {val || '-'}
                                </td>
                              );
                            })}
                            <td className="p-3 border border-slate-100 text-center font-black bg-slate-100/50 text-slate-900">
                              {rowTotal.toLocaleString()}
                            </td>
                          </tr>
                        );
                      })}
                      {/* Column Totals Row */}
                      <tr className="bg-slate-800 font-black text-white border-t-2 border-slate-700">
                        <td colSpan={2} className="p-3 border border-slate-700 sticky left-0 bg-slate-800 z-10 text-right pr-6">
                          TOTAL KESELURUHAN
                        </td>
                        {periodeStats.months.map(m => {
                          let colTotal = 0;
                          periodeStats.insuranceList.forEach(i => {
                            colTotal += (periodeStats.insuranceMatrix[i]?.[m] || 0);
                          });
                          return (
                            <td key={m} className="p-3 border border-slate-700 text-center">
                              {colTotal.toLocaleString()}
                            </td>
                          );
                        })}
                        <td className="p-3 border border-slate-700 text-center bg-slate-900 text-green-400 text-sm">
                          {(() => {
                            let grandTotal = 0;
                            periodeStats.insuranceList.forEach(i => {
                              periodeStats.months.forEach(m => {
                                grandTotal += (periodeStats.insuranceMatrix[i]?.[m] || 0);
                              });
                            });
                            return grandTotal.toLocaleString();
                          })()}
                        </td>
                      </tr>
                    </>
                  )}
                </tbody>
              </table>
            </div>
          </div>

            {/* Age Category Table */}
            <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden mt-8" id="table-periode-age">
              <div className="p-6 border-b border-slate-100 bg-slate-50/50 flex items-center justify-between">
              <div>
                <h3 className="text-lg font-bold text-slate-800">Jumlah Kunjungan Berdasarkan Kategori Umur Per Periode</h3>
                <p className="text-xs text-slate-500">Rekapitulasi kunjungan bulanan per kategori umur</p>
              </div>
              <div className="flex items-center gap-2">
                <button 
                  onClick={() => exportToPNG('table-periode-age', 'kunjungan-umur-periode')}
                  className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                  title="Download PNG"
                >
                  <Download className="w-4 h-4" />
                </button>
                <button 
                  onClick={() => handleExportPeriodeAge('excel')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-colors"
                >
                  <Download className="w-3 h-3" /> Excel
                </button>
                <button 
                  onClick={() => handleExportPeriodeAge('csv')}
                  className="flex items-center gap-1 px-3 py-1.5 bg-slate-100 text-slate-700 rounded-lg text-xs font-bold hover:bg-slate-200 transition-colors"
                >
                  <Download className="w-3 h-3" /> CSV
                </button>
              </div>
            </div>
            
            <div className="overflow-auto max-h-[600px]">
              <table className="w-full border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-800 text-white">
                    <th className="p-3 border border-slate-700 text-center sticky left-0 top-0 bg-slate-800 z-40 w-[40px]">No</th>
                    <th 
                      className="p-3 border border-slate-700 text-left sticky left-[40px] top-0 bg-slate-800 z-30 min-w-[180px] cursor-pointer hover:bg-slate-700 transition-colors"
                      onClick={() => handlePeriodeAgeSort('age')}
                    >
                      <div className="flex items-center justify-between gap-2">
                        <span>Kategori Umur</span>
                        {periodeAgeSort.key === 'age' ? (
                          periodeAgeSort.direction === 'asc' ? <ChevronUp className="w-4 h-4" /> : <ChevronDown className="w-4 h-4" />
                        ) : (
                          <ArrowUpDown className="w-3 h-3 opacity-30" />
                        )}
                      </div>
                    </th>
                    {periodeStats.months.map(m => (
                      <th 
                        key={m} 
                        className="p-3 border border-slate-700 text-center min-w-[90px] cursor-pointer hover:bg-slate-700 transition-colors sticky top-0 bg-slate-800 z-20"
                        onClick={() => handlePeriodeAgeSort(m)}
                      >
                        <div className="flex flex-col items-center gap-1">
                          <span>{periodeStats.monthMap[m]}</span>
                          {periodeAgeSort.key === m ? (
                            periodeAgeSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                          ) : (
                            <ArrowUpDown className="w-2 h-2 opacity-30" />
                          )}
                        </div>
                      </th>
                    ))}
                    <th 
                      className="p-3 border border-slate-700 text-center bg-slate-700 font-bold min-w-[100px] cursor-pointer hover:bg-slate-600 transition-colors sticky top-0 z-20"
                      onClick={() => handlePeriodeAgeSort('total')}
                    >
                      <div className="flex flex-col items-center gap-1">
                        <span>Total</span>
                        {periodeAgeSort.key === 'total' ? (
                          periodeAgeSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                        ) : (
                          <ArrowUpDown className="w-2 h-2 opacity-30" />
                        )}
                      </div>
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {sortedPeriodeAge.length === 0 ? (
                    <tr>
                      <td colSpan={periodeStats.months.length + 3} className="p-8 text-center text-slate-400 italic">
                        Tidak ada data kunjungan untuk periode ini.
                      </td>
                    </tr>
                  ) : (
                    <>
                      {sortedPeriodeAge.map((ageGroup, idx) => {
                        let rowTotal = 0;
                        const label = ageGroupsList.find(g => g.id === ageGroup)?.label || ageGroup;
                        return (
                          <tr key={idx} className="hover:bg-orange-50/50 transition-colors even:bg-slate-50/30">
                            <td className="p-3 border border-slate-100 text-center font-medium text-slate-400 sticky left-0 bg-inherit z-10">
                              {idx + 1}
                            </td>
                            <td className="p-3 border border-slate-100 font-bold text-slate-700 sticky left-[40px] bg-inherit z-10 shadow-[2px_0_5px_rgba(0,0,0,0.03)]">
                              {label}
                            </td>
                            {periodeStats.months.map(m => {
                              const val = periodeStats.ageMatrix[ageGroup]?.[m] || 0;
                              rowTotal += val;
                              return (
                                <td key={m} className={cn(
                                  "p-3 border border-slate-100 text-center",
                                  val > 0 ? "text-orange-600 font-semibold" : "text-slate-300"
                                )}>
                                  {val || '-'}
                                </td>
                              );
                            })}
                            <td className="p-3 border border-slate-100 text-center font-black bg-slate-100/50 text-slate-900">
                              {rowTotal.toLocaleString()}
                            </td>
                          </tr>
                        );
                      })}
                      {/* Column Totals Row */}
                      <tr className="bg-slate-800 font-black text-white border-t-2 border-slate-700">
                        <td colSpan={2} className="p-3 border border-slate-700 sticky left-0 bg-slate-800 z-10 text-right pr-6">
                          TOTAL KESELURUHAN
                        </td>
                        {periodeStats.months.map(m => {
                          let colTotal = 0;
                          periodeStats.ageList.forEach(a => {
                            colTotal += (periodeStats.ageMatrix[a]?.[m] || 0);
                          });
                          return (
                            <td key={m} className="p-3 border border-slate-700 text-center">
                              {colTotal.toLocaleString()}
                            </td>
                          );
                        })}
                        <td className="p-3 border border-slate-700 text-center bg-slate-900 text-orange-400 text-sm">
                          {(() => {
                            let grandTotal = 0;
                            periodeStats.ageList.forEach(a => {
                              periodeStats.months.forEach(m => {
                                grandTotal += (periodeStats.ageMatrix[a]?.[m] || 0);
                              });
                            });
                            return grandTotal.toLocaleString();
                          })()}
                        </td>
                      </tr>
                    </>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
        ) : activeTab === 'programPrioritas' ? (
          <div className="space-y-8">
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
              {/* Chart Program Prioritas */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200" id="chart-priority-program">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <TrendingUp className="w-5 h-5 text-blue-600" />
                    Distribusi Program Prioritas
                  </h3>
                  <button 
                    onClick={() => exportToPNG('chart-priority-program', 'distribusi-program-prioritas')}
                    className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                    title="Download PNG"
                  >
                    <Download className="w-4 h-4" />
                  </button>
                </div>
                <div className="h-[400px] w-full">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart 
                      data={stats?.patientsByPriorityProgram}
                      layout="vertical"
                      margin={{ top: 5, right: 30, left: 100, bottom: 5 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" horizontal={true} vertical={false} stroke="#f1f5f9" />
                      <XAxis type="number" hide />
                      <YAxis 
                        dataKey="name" 
                        type="category" 
                        axisLine={false} 
                        tickLine={false} 
                        tick={{ fontSize: 10, fill: '#64748b', fontWeight: 600 }}
                        width={120}
                      />
                      <Tooltip 
                        cursor={{ fill: '#f8fafc' }}
                        contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                      />
                      <Bar dataKey="value" fill="#3b82f6" radius={[0, 4, 4, 0]} barSize={20}>
                        {showDataLabels && (
                          <LabelList dataKey="value" position="right" style={{ fontSize: 10, fontWeight: 'bold', fill: '#64748b' }} />
                        )}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>

              {/* Table Program Prioritas */}
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200" id="table-priority-program">
                <div className="flex items-center justify-between mb-6">
                  <h3 className="font-bold text-slate-800 flex items-center gap-2">
                    <Users className="w-5 h-5 text-emerald-600" />
                    Data Program Prioritas (Kolom AB)
                  </h3>
                  <button 
                    onClick={() => exportToPNG('table-priority-program', 'data-program-prioritas')}
                    className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                    title="Download PNG"
                  >
                    <Download className="w-4 h-4" />
                  </button>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="border-b border-slate-100">
                        <th className="py-3 px-4 text-[10px] font-bold text-slate-400 uppercase tracking-wider">No</th>
                        <th className="py-3 px-4 text-[10px] font-bold text-slate-400 uppercase tracking-wider">Nama Program</th>
                        <th className="py-3 px-4 text-[10px] font-bold text-slate-400 uppercase tracking-wider text-right">Jumlah Pasien</th>
                        <th className="py-3 px-4 text-[10px] font-bold text-slate-400 uppercase tracking-wider text-right">Persentase</th>
                      </tr>
                    </thead>
                    <tbody>
                      {stats?.patientsByPriorityProgram.map((item, idx) => (
                        <tr key={item.name} className={cn("border-b border-slate-50 hover:bg-slate-50/50 transition-colors", idx % 2 === 0 ? "bg-white" : "bg-slate-50/30")}>
                          <td className="py-3 px-4 text-xs text-slate-500">{idx + 1}</td>
                          <td className="py-3 px-4 text-xs font-bold text-slate-700">{item.name}</td>
                          <td className="py-3 px-4 text-xs text-slate-600 text-right font-medium">{item.value.toLocaleString()}</td>
                          <td className="py-3 px-4 text-xs text-slate-600 text-right">
                            <span className="bg-blue-50 text-blue-600 px-2 py-0.5 rounded-md font-bold text-[10px]">
                              {stats.totalPatients > 0 ? ((item.value / stats.totalPatients) * 100).toFixed(1) : 0}%
                            </span>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot className="bg-slate-50 font-bold">
                      <tr>
                        <td colSpan={2} className="py-3 px-4 text-xs text-slate-700 text-center">TOTAL</td>
                        <td className="py-3 px-4 text-xs text-blue-700 text-right">{stats?.totalPatients.toLocaleString()}</td>
                        <td className="py-3 px-4 text-xs text-blue-700 text-right">100%</td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>
            </div>

            {/* Table Per Periode Program Prioritas */}
            <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200" id="table-priority-program-period">
              <div className="flex items-center justify-between mb-6">
                <div className="flex items-center gap-3">
                  <div className="p-2 bg-orange-50 rounded-lg">
                    <Calendar className="w-5 h-5 text-orange-600" />
                  </div>
                  <div>
                    <h3 className="font-bold text-slate-800">Data Program Prioritas Per Periode</h3>
                    <p className="text-xs text-slate-500">Distribusi program prioritas berdasarkan bulan kunjungan</p>
                  </div>
                </div>
                <button 
                  onClick={() => exportToPNG('table-priority-program-period', 'data-program-prioritas-periode')}
                  className="p-2 hover:bg-slate-100 rounded-lg transition-colors text-slate-400 hover:text-blue-600"
                  title="Download PNG"
                >
                  <Download className="w-4 h-4" />
                </button>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => handleExportPeriodePriorityProgram('excel')}
                    className="flex items-center gap-2 px-3 py-1.5 bg-emerald-50 text-emerald-700 rounded-lg text-xs font-bold hover:bg-emerald-100 transition-colors"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5" /> Excel
                  </button>
                  <button 
                    onClick={() => handleExportPeriodePriorityProgram('csv')}
                    className="flex items-center gap-2 px-3 py-1.5 bg-blue-50 text-blue-700 rounded-lg text-xs font-bold hover:bg-blue-100 transition-colors"
                  >
                    <FileText className="w-3.5 h-3.5" /> CSV
                  </button>
                </div>
              </div>

              <div className="overflow-x-auto rounded-xl border border-slate-100">
                <table className="w-full text-sm border-collapse">
                  <thead>
                    <tr className="bg-slate-800 text-white">
                      <th className="p-3 border border-slate-700 text-left sticky left-0 bg-slate-800 z-20 w-12">No</th>
                      <th 
                        className="p-3 border border-slate-700 text-left min-w-[200px] cursor-pointer hover:bg-slate-700 transition-colors sticky left-[48px] bg-slate-800 z-20"
                        onClick={() => handlePeriodePriorityProgramSort('program')}
                      >
                        <div className="flex items-center gap-2">
                          <span>Nama Program Prioritas</span>
                          {periodePriorityProgramSort.key === 'program' ? (
                            periodePriorityProgramSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                          ) : (
                            <ArrowUpDown className="w-2 h-2 opacity-30" />
                          )}
                        </div>
                      </th>
                      {periodeStats.months.map(m => (
                        <th 
                          key={m} 
                          className="p-3 border border-slate-700 text-center min-w-[90px] cursor-pointer hover:bg-slate-700 transition-colors sticky top-0 bg-slate-800 z-20"
                          onClick={() => handlePeriodePriorityProgramSort(m)}
                        >
                          <div className="flex flex-col items-center gap-1">
                            <span>{periodeStats.monthMap[m]}</span>
                            {periodePriorityProgramSort.key === m ? (
                              periodePriorityProgramSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                            ) : (
                              <ArrowUpDown className="w-2 h-2 opacity-30" />
                            )}
                          </div>
                        </th>
                      ))}
                      <th 
                        className="p-3 border border-slate-700 text-center bg-slate-700 font-bold min-w-[100px] cursor-pointer hover:bg-slate-600 transition-colors sticky top-0 z-20"
                        onClick={() => handlePeriodePriorityProgramSort('total')}
                      >
                        <div className="flex flex-col items-center gap-1">
                          <span>Total</span>
                          {periodePriorityProgramSort.key === 'total' ? (
                            periodePriorityProgramSort.direction === 'asc' ? <ChevronUp className="w-3 h-3" /> : <ChevronDown className="w-3 h-3" />
                          ) : (
                            <ArrowUpDown className="w-2 h-2 opacity-30" />
                          )}
                        </div>
                      </th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {sortedPeriodePriorityProgram.length === 0 ? (
                      <tr>
                        <td colSpan={periodeStats.months.length + 3} className="p-8 text-center text-slate-400 italic">
                          Tidak ada data program prioritas untuk periode ini.
                        </td>
                      </tr>
                    ) : (
                      <>
                        {sortedPeriodePriorityProgram.map((program, idx) => {
                          let rowTotal = 0;
                          return (
                            <tr key={idx} className="hover:bg-blue-50/50 transition-colors even:bg-slate-50/30">
                              <td className="p-3 border border-slate-100 text-center font-medium text-slate-400 sticky left-0 bg-inherit z-10">
                                {idx + 1}
                              </td>
                              <td className="p-3 border border-slate-100 font-bold text-slate-700 sticky left-[48px] bg-inherit z-10 shadow-[2px_0_5px_rgba(0,0,0,0.03)]">
                                {program}
                              </td>
                              {periodeStats.months.map(m => {
                                const val = periodeStats.priorityProgramMatrix[program]?.[m] || 0;
                                rowTotal += val;
                                return (
                                  <td key={m} className={cn(
                                    "p-3 border border-slate-100 text-center",
                                    val > 0 ? "text-blue-600 font-semibold" : "text-slate-300"
                                  )}>
                                    {val || '-'}
                                  </td>
                                );
                              })}
                              <td className="p-3 border border-slate-100 text-center font-black bg-slate-100/50 text-slate-900">
                                {rowTotal.toLocaleString()}
                              </td>
                            </tr>
                          );
                        })}
                        {/* Column Totals Row */}
                        <tr className="bg-slate-800 font-black text-white border-t-2 border-slate-700">
                          <td colSpan={2} className="p-3 border border-slate-700 sticky left-0 bg-slate-800 z-10 text-right pr-6">
                            TOTAL KESELURUHAN
                          </td>
                          {periodeStats.months.map(m => {
                            let colTotal = 0;
                            periodeStats.priorityProgramList.forEach(p => {
                              colTotal += (periodeStats.priorityProgramMatrix[p]?.[m] || 0);
                            });
                            return (
                              <td key={m} className="p-3 border border-slate-700 text-center">
                                {colTotal.toLocaleString()}
                              </td>
                            );
                          })}
                          <td className="p-3 border border-slate-700 text-center bg-slate-900 text-orange-400 text-sm">
                            {(() => {
                              let grandTotal = 0;
                              periodeStats.priorityProgramList.forEach(p => {
                                periodeStats.months.forEach(m => {
                                  grandTotal += (periodeStats.priorityProgramMatrix[p]?.[m] || 0);
                                });
                              });
                              return grandTotal.toLocaleString();
                            })()}
                          </td>
                        </tr>
                      </>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            {/* Left: Search Section */}
            <div className="lg:col-span-5 space-y-6">
              <button 
                onClick={() => setActiveTab('diagnosa')}
                className="flex items-center gap-2 text-slate-500 hover:text-blue-600 font-medium text-sm transition-colors mb-4"
              >
                <ArrowLeft className="w-4 h-4" /> Kembali RL 51 Mobiditas Pasien Rawat Jalan
              </button>
              
              <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200">
                <h3 className="text-lg font-bold text-slate-800 mb-4">Search Nama Penyakit</h3>
                <div className="space-y-4">
                  <div className="relative">
                    <input 
                      type="text" 
                      placeholder="Search Nama Penyakit / KODE ICD10" 
                      className="w-full pl-4 pr-4 py-3 bg-slate-50 border-transparent focus:bg-white focus:ring-2 focus:ring-blue-500 rounded-xl text-sm transition-all"
                      value={inputSearchTerm}
                      onChange={(e) => setInputSearchTerm(e.target.value)}
                    />
                  </div>
                  <button className="bg-emerald-50 text-emerald-700 font-bold px-6 py-2 rounded-lg flex items-center gap-2 hover:bg-emerald-100 transition-colors">
                    <Search className="w-4 h-4" /> Cari
                  </button>
                </div>

                <div className="mt-6 max-h-[600px] overflow-auto border border-slate-100 rounded-xl">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-slate-50 border-b border-slate-100">
                        <th className="p-3 text-left font-semibold text-slate-500 w-12 text-[11px] uppercase tracking-wider sticky top-0 bg-slate-50 z-10 shadow-[inset_0_-1px_0_rgba(241,245,249,1)]">No.</th>
                        <th className="p-3 text-left font-semibold text-slate-500 text-[11px] uppercase tracking-wider sticky top-0 bg-slate-50 z-10 shadow-[inset_0_-1px_0_rgba(241,245,249,1)]">Code ICD 10</th>
                        <th className="p-3 text-left font-semibold text-slate-500 text-[11px] uppercase tracking-wider sticky top-0 bg-slate-50 z-10 shadow-[inset_0_-1px_0_rgba(241,245,249,1)]">Deskripsi ICD 10</th>
                        <th className="p-3 text-center font-semibold text-slate-500 text-[11px] uppercase tracking-wider sticky top-0 bg-slate-50 z-10 shadow-[inset_0_-1px_0_rgba(241,245,249,1)]">Action</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-50">
                      {filteredInputDiseases.map((row, i) => (
                        <tr key={i} className="hover:bg-slate-50 transition-colors">
                          <td className="p-3 text-slate-500">{i + 1}</td>
                          <td className="p-3 font-mono font-medium text-slate-700">{row.icd}</td>
                          <td className="p-3 text-slate-600">{row.diagnosis}</td>
                          <td className="p-3 text-center">
                            <button 
                              onClick={() => handleSelectDisease(row)}
                              className={cn(
                                "px-3 py-1 border rounded-lg text-xs font-bold transition-all",
                                selectedDiseaseForInput?.icd === row.icd 
                                  ? "bg-emerald-600 border-emerald-600 text-white" 
                                  : "border-emerald-500 text-emerald-600 hover:bg-emerald-50"
                              )}
                            >
                              {selectedDiseaseForInput?.icd === row.icd ? 'Terpilih' : 'Pilih & Lihat Data'}
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Right: Input Section */}
            <div className="lg:col-span-7">
              {selectedDiseaseForInput ? (
                <div className="bg-white p-8 rounded-2xl shadow-sm border border-slate-200">
                  <div className="flex items-center justify-between mb-8">
                    <div>
                      <h3 className="text-xl font-bold text-slate-800">Periode Laporan</h3>
                      <p className="text-slate-500">Tambah Data Penyakit {selectedDiseaseForInput.diagnosis}</p>
                      <div className="flex flex-wrap gap-4 mt-2">
                        <span className="text-[10px] bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full font-bold">
                          Total Kunjungan: {selectedDiseaseForInput.totalVisits?.total || 0}
                        </span>
                        <span className="text-[10px] bg-emerald-100 text-emerald-700 px-2 py-0.5 rounded-full font-bold">
                          Total Kasus Baru: {selectedDiseaseForInput.totalNew?.total || 0}
                        </span>
                        <span className="text-[10px] text-slate-400 italic flex items-center gap-1">
                          <AlertCircle className="w-3 h-3" /> Data di bawah otomatis terisi dari spreadsheet saat ini
                        </span>
                      </div>
                    </div>
                    <div className="flex items-center gap-2">
                      <button 
                        onClick={() => {
                          const cleared: Record<string, { L: number, P: number }> = {};
                          ageGroupsList.forEach(g => cleared[g.id] = { L: 0, P: 0 });
                          setInputValues(cleared);
                          setTotalVisitsInput({ L: 0, P: 0 });
                        }}
                        className="text-xs font-bold text-slate-500 hover:text-slate-700 flex items-center gap-1 bg-slate-100 px-3 py-2 rounded-lg"
                      >
                        Clear All
                      </button>
                      <button 
                        onClick={() => handleSelectDisease(selectedDiseaseForInput)}
                        className="text-xs font-bold text-blue-600 hover:text-blue-700 flex items-center gap-1 bg-blue-50 px-3 py-2 rounded-lg"
                      >
                        <RefreshCw className="w-3 h-3" /> Reset ke Data Diagnosa
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 gap-6 mb-8">
                    <div>
                      <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Tahun</label>
                      <input 
                        type="text" 
                        className="w-full px-4 py-3 bg-slate-50 border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500"
                        value={inputYear}
                        onChange={(e) => setInputYear(e.target.value)}
                      />
                    </div>
                    <div>
                      <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider mb-2">Bulan</label>
                      <select 
                        className="w-full px-4 py-3 bg-slate-50 border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500"
                        value={inputMonth}
                        onChange={(e) => setInputMonth(e.target.value)}
                      >
                        {['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'].map(m => (
                          <option key={m} value={m}>{m}</option>
                        ))}
                      </select>
                    </div>
                  </div>

                  <div className="border border-slate-100 rounded-2xl overflow-hidden mb-8">
                    <div className="max-h-[500px] overflow-y-auto">
                      <table className="w-full text-sm">
                        <thead className="sticky top-0 bg-slate-50 z-10">
                          <tr className="border-b border-slate-100">
                            <th className="p-4 text-left font-semibold text-slate-500 w-12">No.</th>
                            <th className="p-4 text-left font-semibold text-slate-500">Golongan Berdasarkan Umur</th>
                            <th className="p-4 text-center font-semibold text-slate-500">Laki Laki</th>
                            <th className="p-4 text-center font-semibold text-slate-500">Perempuan</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-50">
                          {ageGroupsList.map((group, i) => (
                            <tr key={group.id}>
                              <td className="p-4 text-slate-400 font-medium">{i + 1}</td>
                              <td className="p-4 font-medium text-slate-700">Umur {group.label}</td>
                              <td className="p-4">
                                <input 
                                  type="number" 
                                  className={cn(
                                    "w-full px-3 py-2 border rounded-lg text-center transition-all focus:ring-2 focus:ring-blue-500",
                                    (inputValues[group.id]?.L || 0) > 0 
                                      ? "bg-blue-50 border-blue-200 text-blue-700 font-bold" 
                                      : "bg-slate-50 border-slate-200"
                                  )}
                                  value={inputValues[group.id]?.L || 0}
                                  onChange={(e) => handleInputChange(group.id, 'L', e.target.value)}
                                />
                              </td>
                              <td className="p-4">
                                <input 
                                  type="number" 
                                  className={cn(
                                    "w-full px-3 py-2 border rounded-lg text-center transition-all focus:ring-2 focus:ring-blue-500",
                                    (inputValues[group.id]?.P || 0) > 0 
                                      ? "bg-pink-50 border-pink-200 text-pink-700 font-bold" 
                                      : "bg-slate-50 border-slate-200"
                                  )}
                                  value={inputValues[group.id]?.P || 0}
                                  onChange={(e) => handleInputChange(group.id, 'P', e.target.value)}
                                />
                              </td>
                            </tr>
                          ))}
                          <tr className="bg-slate-100 font-bold text-slate-900">
                            <td className="p-4 text-slate-400">26</td>
                            <td className="p-4">TOTAL KASUS BARU (RL 5.1)</td>
                            <td className="p-4 text-center text-blue-700">
                              {Object.values(inputValues).reduce((acc, curr) => acc + curr.L, 0)}
                            </td>
                            <td className="p-4 text-center text-pink-700">
                              {Object.values(inputValues).reduce((acc, curr) => acc + curr.P, 0)}
                            </td>
                          </tr>
                          <tr className="bg-blue-50 font-bold text-slate-900 border-t-2 border-blue-100">
                            <td className="p-4 text-slate-400">27</td>
                            <td className="p-4">TOTAL JUMLAH KUNJUNGAN PASIEN</td>
                            <td className="p-4">
                              <input 
                                type="number" 
                                className="w-full px-3 py-2 border border-blue-200 rounded-lg text-center bg-white focus:ring-2 focus:ring-blue-500 text-blue-700"
                                value={totalVisitsInput.L}
                                onChange={(e) => setTotalVisitsInput(prev => ({ ...prev, L: parseInt(e.target.value) || 0 }))}
                              />
                            </td>
                            <td className="p-4">
                              <input 
                                type="number" 
                                className="w-full px-3 py-2 border border-pink-200 rounded-lg text-center bg-white focus:ring-2 focus:ring-blue-500 text-pink-700"
                                value={totalVisitsInput.P}
                                onChange={(e) => setTotalVisitsInput(prev => ({ ...prev, P: parseInt(e.target.value) || 0 }))}
                              />
                            </td>
                          </tr>
                          <tr className="bg-slate-200 font-black text-slate-900">
                            <td className="p-4"></td>
                            <td className="p-4">TOTAL KESELURUHAN KUNJUNGAN</td>
                            <td colSpan={2} className="p-4 text-center text-lg">
                              {totalVisitsInput.L + totalVisitsInput.P}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                  <button className="bg-emerald-50 text-emerald-700 font-bold px-8 py-3 rounded-xl flex items-center gap-2 hover:bg-emerald-100 transition-colors">
                    <Save className="w-5 h-5" /> Simpan
                  </button>
                </div>
              ) : (
                <div className="bg-white p-12 rounded-2xl shadow-sm border border-slate-200 flex flex-col items-center justify-center text-center">
                  <div className="bg-slate-50 p-6 rounded-full mb-4">
                    <Plus className="w-12 h-12 text-slate-300" />
                  </div>
                  <h3 className="text-lg font-bold text-slate-800 mb-2">Pilih Penyakit</h3>
                  <p className="text-slate-500 max-w-xs">Pilih penyakit dari daftar di sebelah kiri untuk mulai menginput data RL 5.1</p>
                </div>
              )}
            </div>
          </div>
        )}
      </main>

      {/* Footer */}
      <footer className="px-4 sm:px-6 lg:px-8 mt-12">
        <div className="border-t border-slate-200 pt-8 flex flex-col md:flex-row items-center justify-between gap-4">
          <p className="text-sm text-slate-400">© 2026 Dashboard Poliklinik. Data diperbarui secara real-time dari Google Sheets.</p>
          <div className="flex items-center gap-6">
            <a href={SHEET_URL} target="_blank" rel="noopener noreferrer" className="text-xs font-bold text-blue-600 hover:text-blue-700 flex items-center gap-1">
              <Download className="w-3 h-3" /> Sumber Data (CSV)
            </a>
          </div>
        </div>
      </footer>
    </div>
  );
}
