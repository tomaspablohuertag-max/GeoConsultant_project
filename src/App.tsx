/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useRef, useMemo } from 'react';
import { APIProvider, useMapsLibrary } from '@vis.gl/react-google-maps';
import { motion, AnimatePresence } from 'motion/react';
import {
  Calculator,
  FileSpreadsheet,
  Plus,
  Trash2,
  ArrowRightLeft,
  Navigation,
  CheckCircle2,
  AlertCircle,
  Cloud,
  CloudOff,
  RefreshCw,
  Search,
  Edit2,
  X,
  Loader2,
  RotateCcw,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { cn } from './lib/utils';

// --- Constants & Types ---

const API_KEY = process.env.GOOGLE_MAPS_PLATFORM_KEY || '';
const hasValidKey = Boolean(API_KEY) && API_KEY !== 'YOUR_API_KEY';

interface MileageRecord {
  id: string;
  consultantName: string;
  origin: string;
  destination: string;
  days: number;
  distanceKm: number;
  totalDistanceKm: number;
  date: string; // DD/MM/YYYY
  isRoundTrip: boolean;
  syncedToOneDrive?: boolean;
}

interface Filters {
  consultant: string;
  dateFrom: string;
  dateTo: string;
}

// --- Helpers ---

function getTodayInputDate(): string {
  const d = new Date();
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function inputDateToDisplay(value: string): string {
  if (!value) return '';
  const [year, month, day] = value.split('-');
  return `${day}/${month}/${year}`;
}

function displayDateToInput(value: string): string {
  if (!value) return '';
  const [day, month, year] = value.split('/');
  if (!day || !month || !year) return '';
  return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
}

function parseGoogleDistanceError(status: string): string {
  switch (status) {
    case 'NOT_FOUND':
      return 'No se encontró una de las direcciones. Verifica origen y destino.';
    case 'ZERO_RESULTS':
      return 'No hay ruta en coche disponible entre estas direcciones.';
    case 'MAX_ROUTE_LENGTH_EXCEEDED':
      return 'La ruta es demasiado larga.';
    case 'OVER_QUERY_LIMIT':
      return 'Se excedió el límite de consultas de Google Maps. Intenta más tarde.';
    case 'REQUEST_DENIED':
      return 'Acceso denegado a Google Maps. Revisa tu clave de API.';
    case 'UNKNOWN_ERROR':
    default:
      return 'Error al calcular la distancia. Intenta de nuevo.';
  }
}

// --- Components ---

function ApiKeySplashScreen() {
  return (
    <div className="flex flex-col items-center justify-center min-h-screen bg-[#fcfcfc] p-6 font-sans">
      <motion.div
        initial={{ opacity: 0, y: 10 }}
        animate={{ opacity: 1, y: 0 }}
        className="max-w-md w-full bg-white rounded-2xl border border-slate-200 p-8 shadow-sm"
      >
        <div className="w-12 h-12 bg-blue-600 rounded-xl flex items-center justify-center mb-6">
          <Navigation className="text-white w-6 h-6" />
        </div>
        <h2 className="text-xl font-bold text-slate-900 mb-2">Configuración Requerida</h2>
        <p className="text-sm text-slate-500 mb-8">
          Se necesita una clave de API de Google Maps para calcular las distancias con exactitud.
        </p>

        <div className="space-y-6">
          <div className="flex gap-4">
            <div className="w-6 h-6 rounded-full bg-slate-100 flex items-center justify-center text-[10px] font-bold text-slate-500 shrink-0">1</div>
            <div>
              <p className="text-sm font-medium text-slate-800">Obtén tu clave de API</p>
              <p className="text-xs text-slate-400 mt-1">
                Visita <a href="https://console.cloud.google.com/google/maps-apis/start" target="_blank" rel="noopener noreferrer" className="text-blue-600 hover:underline">Google Cloud Console</a>.
              </p>
            </div>
          </div>
          <div className="flex gap-4">
            <div className="w-6 h-6 rounded-full bg-slate-100 flex items-center justify-center text-[10px] font-bold text-slate-500 shrink-0">2</div>
            <div>
              <p className="text-sm font-medium text-slate-800">Configura la variable de entorno</p>
              <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                Añade <code>GOOGLE_MAPS_PLATFORM_KEY</code> en el archivo <code>.env</code> de la raíz del proyecto y reinicia el servidor.
              </p>
            </div>
          </div>
        </div>
      </motion.div>
    </div>
  );
}

function LoadingSpinner({ className }: { className?: string }) {
  return <Loader2 className={cn('animate-spin', className)} />;
}

function ErrorBanner({ message, onClose }: { message: string; onClose: () => void }) {
  return (
    <motion.div
      initial={{ opacity: 0, y: -8 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -8 }}
      className="flex items-start gap-2.5 bg-red-50 border border-red-200 text-red-700 rounded-lg p-3 text-xs"
    >
      <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
      <span className="flex-1 leading-relaxed">{message}</span>
      <button onClick={onClose} className="text-red-400 hover:text-red-600">
        <X className="w-4 h-4" />
      </button>
    </motion.div>
  );
}

function SuccessBanner({ message }: { message: string }) {
  return (
    <motion.div
      initial={{ opacity: 0, y: -8 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -8 }}
      className="flex items-center gap-2.5 bg-green-50 border border-green-200 text-green-700 rounded-lg p-3 text-xs"
    >
      <CheckCircle2 className="w-4 h-4 shrink-0" />
      <span className="leading-relaxed">{message}</span>
    </motion.div>
  );
}

function AddressAutocomplete({
  label,
  placeholder,
  value,
  onChange,
  disabled,
}: {
  label: string;
  placeholder: string;
  value: string;
  onChange: (val: string) => void;
  disabled?: boolean;
}) {
  const [inputValue, setInputValue] = useState(value);
  const [isReady, setIsReady] = useState(false);
  const placesLib = useMapsLibrary('places');
  const inputRef = useRef<HTMLInputElement>(null);
  const autocompleteRef = useRef<google.maps.places.Autocomplete | null>(null);

  useEffect(() => {
    setInputValue(value);
  }, [value]);

  useEffect(() => {
    if (!placesLib || !inputRef.current) return;

    const BARCELONA_BOUNDS = {
      north: 41.8,
      south: 41.0,
      east: 2.6,
      west: 1.6,
    };

    const options = {
      fields: ['formatted_address', 'geometry', 'name'],
      bounds: BARCELONA_BOUNDS,
      componentRestrictions: { country: 'es' },
      types: ['establishment', 'geocode'],
    };

    autocompleteRef.current = new placesLib.Autocomplete(inputRef.current, options);
    autocompleteRef.current.addListener('place_changed', () => {
      const place = autocompleteRef.current?.getPlace();
      if (place?.formatted_address) {
        onChange(place.formatted_address);
        setInputValue(place.formatted_address);
      } else if (place?.name) {
        onChange(place.name);
        setInputValue(place.name);
      }
    });
    setIsReady(true);

    return () => {
      if (autocompleteRef.current) google.maps.event.clearInstanceListeners(autocompleteRef.current);
    };
  }, [placesLib, onChange]);

  return (
    <div>
      <label className="block text-[11px] font-semibold text-slate-500 mb-1.5 uppercase tracking-wide">
        {label}
      </label>
      <div className="relative">
        <input
          ref={inputRef}
          type="text"
          value={inputValue}
          disabled={disabled}
          onChange={(e) => {
            setInputValue(e.target.value);
            onChange(e.target.value);
          }}
          placeholder={placeholder}
          className="w-full border border-slate-200 rounded-lg px-3 py-2.5 text-sm placeholder:text-slate-300 outline-none focus:ring-2 focus:ring-blue-100 transition-all bg-white disabled:opacity-60"
        />
        {!isReady && (
          <div className="absolute right-3 top-1/2 -translate-y-1/2">
            <LoadingSpinner className="w-4 h-4 text-slate-300" />
          </div>
        )}
      </div>
    </div>
  );
}

// --- Main App Logic ---

export default function App() {
  if (!hasValidKey) {
    return <ApiKeySplashScreen />;
  }

  return (
    <APIProvider apiKey={API_KEY} version="weekly">
      <MileageApp />
    </APIProvider>
  );
}

function MileageApp() {
  // Form state
  const [consultantName, setConsultantName] = useState('');
  const [origin, setOrigin] = useState('');
  const [destination, setDestination] = useState('');
  const [days, setDays] = useState(1);
  const [recordDate, setRecordDate] = useState(getTodayInputDate());
  const [isRoundTrip, setIsRoundTrip] = useState(true);
  const [editingId, setEditingId] = useState<string | null>(null);

  // Preview
  const [previewDistance, setPreviewDistance] = useState<number | null>(null);

  // Records & UI
  const [records, setRecords] = useState<MileageRecord[]>([]);
  const [isLoadingRecords, setIsLoadingRecords] = useState(true);
  const [isCalculating, setIsCalculating] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  // Filters
  const [filters, setFilters] = useState<Filters>(({
    consultant: '',
    dateFrom: '',
    dateTo: '',
  }));

  // OneDrive Integration States
  const [onedriveStatus, setOnedriveStatus] = useState<{
    connected: boolean;
    isConfigured: boolean;
    userName?: string;
    userEmail?: string;
  }>({ connected: false, isConfigured: false });
  const [isSyncing, setIsSyncing] = useState(false);

  useEffect(() => {
    fetchStatus();
    fetchRecords();
  }, []);

  // Auto-hide success messages
  useEffect(() => {
    if (!success) return;
    const timer = setTimeout(() => setSuccess(null), 4000);
    return () => clearTimeout(timer);
  }, [success]);

  // Invalidate distance preview when route parameters change
  useEffect(() => {
    setPreviewDistance(null);
  }, [origin, destination, days, isRoundTrip]);

  const fetchStatus = async () => {
    try {
      const res = await fetch('/api/auth/status');
      if (res.ok) {
        const data = await res.json();
        setOnedriveStatus(data);
      }
    } catch (err) {
      console.error('Error fetching OneDrive status:', err);
    }
  };

  const fetchRecords = async () => {
    setIsLoadingRecords(true);
    try {
      const res = await fetch('/api/records');
      if (res.ok) {
        const data = await res.json();
        setRecords(data);
      } else {
        setError('No se pudieron cargar los registros.');
      }
    } catch (err) {
      setError('Error de red al cargar los registros.');
      console.error('Error fetching records:', err);
    } finally {
      setIsLoadingRecords(false);
    }
  };

  useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      const origin = event.origin;
      if (!origin.endsWith('.run.app') && !origin.includes('localhost')) {
        return;
      }
      if (event.data?.type === 'MS_AUTH_SUCCESS') {
        fetchStatus();
        fetchRecords();
        fetch('/api/sync', { method: 'POST' }).then(() => {
          fetchRecords();
          fetchStatus();
        });
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const handleConnectOneDrive = async () => {
    try {
      const res = await fetch('/api/auth/microsoft/url');
      if (!res.ok) {
        const errData = await res.json();
        setError(errData.error || 'No se pudo obtener la URL de conexión.');
        return;
      }
      const { url } = await res.json();
      const popup = window.open(url, 'OneDriveAuth', 'width=600,height=700');
      if (!popup) {
        setError('Por favor habilita las ventanas emergentes para conectar OneDrive.');
      }
    } catch (err) {
      setError('Error al iniciar la conexión con Microsoft.');
    }
  };

  const handleLogoutOneDrive = async () => {
    try {
      const res = await fetch('/api/auth/logout', { method: 'POST' });
      if (res.ok) {
        fetchStatus();
        fetchRecords();
      }
    } catch (err) {
      console.error('Error disconnecting OneDrive:', err);
    }
  };

  const handleBulkSync = async () => {
    setIsSyncing(true);
    setError(null);
    try {
      const res = await fetch('/api/sync', { method: 'POST' });
      const data = await res.json();
      if (res.ok) {
        fetchRecords();
        setSuccess(`Sincronizados ${data.syncedCount} de ${data.totalCount} registros.`);
      } else {
        setError(data.error || 'Error de sincronización.');
      }
    } catch (err) {
      setError('Error de red durante la sincronización.');
    } finally {
      setIsSyncing(false);
    }
  };

  const validateForm = (): string | null => {
    if (!consultantName.trim()) return 'El nombre del consultor es obligatorio.';
    if (!origin.trim()) return 'El origen es obligatorio.';
    if (!destination.trim()) return 'El destino es obligatorio.';
    if (origin.trim().toLowerCase() === destination.trim().toLowerCase()) {
      return 'El origen y el destino deben ser diferentes.';
    }
    if (days < 1) return 'El número de días debe ser al menos 1.';
    if (!recordDate) return 'La fecha es obligatoria.';
    return null;
  };

  const calculateDistance = async (): Promise<number | null> => {
    if (!window.google || !origin || !destination) return null;
    const service = new google.maps.DistanceMatrixService();
    try {
      const response = await service.getDistanceMatrix({
        origins: [origin],
        destinations: [destination],
        travelMode: google.maps.TravelMode.DRIVING,
        unitSystem: google.maps.UnitSystem.METRIC,
      });
      const element = response.rows[0].elements[0];
      if (element.status === 'OK') return element.distance.value / 1000;
      throw new Error(parseGoogleDistanceError(element.status));
    } catch (err: any) {
      setError(err.message || 'Error al calcular la distancia.');
      return null;
    }
  };

  const handleCalculate = async () => {
    const validationError = validateForm();
    if (validationError) {
      setError(validationError);
      return;
    }
    setIsCalculating(true);
    setError(null);
    setPreviewDistance(null);
    const distance = await calculateDistance();
    if (distance !== null) {
      setPreviewDistance(distance);
    }
    setIsCalculating(false);
  };

  const handleSave = async () => {
    const validationError = validateForm();
    if (validationError) {
      setError(validationError);
      return;
    }

    let distance = previewDistance;
    if (distance === null) {
      setIsSaving(true);
      distance = await calculateDistance();
      setIsSaving(false);
    }
    if (distance === null) return;

    const multiplier = isRoundTrip ? 2 : 1;
    const total = Number((distance * multiplier * days).toFixed(2));
    const payload = {
      id: editingId || undefined,
      consultantName: consultantName.trim(),
      origin: origin.trim(),
      destination: destination.trim(),
      days,
      distanceKm: Number(distance.toFixed(2)),
      totalDistanceKm: total,
      date: inputDateToDisplay(recordDate),
      isRoundTrip,
    };

    setIsSaving(true);
    setError(null);
    try {
      const url = editingId ? `/api/records/${editingId}` : '/api/records';
      const method = editingId ? 'PUT' : 'POST';
      const response = await fetch(url, {
        method,
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });

      if (response.ok) {
        const saved = await response.json();
        if (editingId) {
          setRecords((prev) => prev.map((r) => (r.id === saved.id ? saved : r)));
          setSuccess('Registro actualizado correctamente.');
        } else {
          setRecords((prev) => [saved, ...prev]);
          setSuccess('Registro añadido correctamente.');
        }
        resetForm();
      } else {
        const errData = await response.json();
        setError(errData.error || 'Error al guardar el registro.');
      }
    } catch (err) {
      setError('Error al comunicarse con el servidor.');
    } finally {
      setIsSaving(false);
    }
  };

  const resetForm = () => {
    setConsultantName('');
    setOrigin('');
    setDestination('');
    setDays(1);
    setRecordDate(getTodayInputDate());
    setIsRoundTrip(true);
    setEditingId(null);
    setPreviewDistance(null);
  };

  const handleEdit = (record: MileageRecord) => {
    setConsultantName(record.consultantName);
    setOrigin(record.origin);
    setDestination(record.destination);
    setDays(record.days);
    setRecordDate(displayDateToInput(record.date) || getTodayInputDate());
    setIsRoundTrip(record.isRoundTrip ?? true);
    setEditingId(record.id);
    setPreviewDistance(record.distanceKm);
    setError(null);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleDelete = async (id: string) => {
    const record = records.find((r) => r.id === id);
    if (!record) return;
    const confirmed = window.confirm(
      `¿Eliminar el registro de ${record.consultantName}\n${record.origin} → ${record.destination}?`
    );
    if (!confirmed) return;

    try {
      const res = await fetch(`/api/records/${id}`, { method: 'DELETE' });
      if (res.ok) {
        setRecords((prev) => prev.filter((r) => r.id !== id));
        setSuccess('Registro eliminado.');
      } else {
        setError('Error al eliminar el registro del servidor.');
      }
    } catch (err) {
      setError('Error de red al eliminar el registro.');
    }
  };

  const cancelEdit = () => {
    resetForm();
  };

  const filteredRecords = useMemo(() => {
    return records.filter((r) => {
      const matchesConsultant = r.consultantName.toLowerCase().includes(filters.consultant.toLowerCase());
      const inputDate = displayDateToInput(r.date);
      const matchesFrom = !filters.dateFrom || (inputDate && inputDate >= filters.dateFrom);
      const matchesTo = !filters.dateTo || (inputDate && inputDate <= filters.dateTo);
      return matchesConsultant && matchesFrom && matchesTo;
    });
  }, [records, filters]);

  const totalKm = useMemo(
    () => filteredRecords.reduce((acc, r) => acc + r.totalDistanceKm, 0),
    [filteredRecords]
  );

  const exportToExcel = () => {
    if (filteredRecords.length === 0) return;
    const logData = filteredRecords.map((r) => ({
      Consultor: r.consultantName,
      Origen: r.origin,
      Destino: r.destination,
      Días: r.days,
      'Distancia Trayecto (Km)': r.distanceKm,
      'Total Trayecto (Ida y Vta)': r.isRoundTrip ? r.distanceKm * 2 : r.distanceKm,
      'Total Acumulado (Km)': r.totalDistanceKm,
      Fecha: r.date,
    }));
    const summaryMap = filteredRecords.reduce((acc, r) => {
      if (!acc[r.consultantName]) {
        acc[r.consultantName] = {
          Consultor: r.consultantName,
          'Total Viajes': 0,
          'Total Días': 0,
          'Km Totales Acumulados': 0,
        };
      }
      acc[r.consultantName]['Total Viajes'] += 1;
      acc[r.consultantName]['Total Días'] += r.days;
      acc[r.consultantName]['Km Totales Acumulados'] += r.totalDistanceKm;
      return acc;
    }, {} as Record<string, any>);
    const summaryData = Object.values(summaryMap);
    const workbook = XLSX.utils.book_new();
    const logSheet = XLSX.utils.json_to_sheet(logData);
    const summarySheet = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Resumen por Consultor');
    XLSX.utils.book_append_sheet(workbook, logSheet, 'Detalle de Trayectos');
    XLSX.writeFile(workbook, `Reporte_Kilometraje_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div className="min-h-screen w-full bg-[#fcfcfc] flex flex-col font-sans text-slate-800">
      {/* Navbar */}
      <nav className="h-16 px-4 sm:px-8 border-b border-slate-100 flex items-center justify-between bg-white shrink-0">
        <div className="flex items-center gap-3">
          <div className="w-8 h-8 bg-blue-600 rounded-lg flex items-center justify-center">
            <svg className="w-5 h-5 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 20l-5.447-2.724A1 1 0 013 16.382V5.618a1 1 0 011.447-.894L9 7m0 13l6-3m-6 3V7m6 10l4.553 2.2276A1 1 0 0021 18.382V7.618a1 1 0 00-.553-.894L15 4m0 13V4m0 0L9 7"></path>
            </svg>
          </div>
          <span className="text-xl font-bold tracking-tight text-slate-900">
            GeoConsult <span className="font-light text-slate-400">Pro</span>
          </span>
        </div>
        <div className="flex items-center gap-4 sm:gap-6">
          <div className="hidden sm:flex flex-col items-end">
            <span className="text-sm font-medium">Gestor de Trayectos</span>
            <span className="text-[10px] uppercase tracking-widest text-slate-400">Panel Consultor</span>
          </div>
          <div className="w-10 h-10 rounded-full bg-slate-100 border border-slate-200"></div>
        </div>
      </nav>

      <main className="flex-1 p-4 sm:p-6 flex flex-col xl:flex-row gap-6 overflow-hidden">
        {/* Sidebar */}
        <div className="w-full xl:w-80 flex flex-col gap-6 shrink-0 overflow-y-auto">
          <section className="bg-white rounded-2xl border border-slate-200 p-6 flex flex-col gap-5 shadow-sm">
            <div className="flex items-center justify-between">
              <h2 className="text-xs font-bold uppercase tracking-widest text-slate-400">{editingId ? 'Editar Entrada' : 'Nueva Entrada'}</h2>
              {editingId && (
                <button onClick={cancelEdit} className="text-[10px] text-slate-400 hover:text-slate-600 flex items-center gap-1">
                  <RotateCcw className="w-3 h-3" /> Cancelar
                </button>
              )}
            </div>

            <div className="space-y-4">
              <div>
                <label className="block text-[11px] font-semibold text-slate-500 mb-1.5 uppercase tracking-wide">
                  Consultor
                </label>
                <input
                  type="text"
                  value={consultantName}
                  onChange={(e) => setConsultantName(e.target.value)}
                  placeholder="Nombre del consultor"
                  className="w-full border border-slate-200 rounded-lg px-3 py-2 text-sm placeholder:text-slate-300 outline-none focus:ring-2 focus:ring-blue-100 bg-white"
                />
              </div>

              <AddressAutocomplete
                label="Origen (Domicilio)"
                placeholder="Calle Serrano 12, Madrid"
                value={origin}
                onChange={setOrigin}
              />

              <AddressAutocomplete
                label="Destino (Planta o Cliente)"
                placeholder="Ej: Novartis, Grifols, Sanofi..."
                value={destination}
                onChange={setDestination}
              />

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-[11px] font-semibold text-slate-500 mb-1.5 uppercase tracking-wide">
                    Fecha
                  </label>
                  <input
                    type="date"
                    value={recordDate}
                    onChange={(e) => setRecordDate(e.target.value)}
                    className="w-full border border-slate-200 rounded-lg px-3 py-2.5 text-sm outline-none focus:ring-2 focus:ring-blue-100 bg-white"
                  />
                </div>
                <div>
                  <label className="block text-[11px] font-semibold text-slate-500 mb-1.5 uppercase tracking-wide">
                    Días
                  </label>
                  <input
                    type="number"
                    min={1}
                    value={days}
                    onChange={(e) => setDays(Math.max(1, parseInt(e.target.value) || 1))}
                    className="w-full border border-slate-200 rounded-lg px-3 py-2.5 text-sm outline-none focus:ring-2 focus:ring-blue-100 bg-white"
                  />
                </div>
              </div>

              <label className="flex items-center gap-2 cursor-pointer select-none">
                <input
                  type="checkbox"
                  checked={isRoundTrip}
                  onChange={(e) => setIsRoundTrip(e.target.checked)}
                  className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-100"
                />
                <span className="text-xs text-slate-600">Aplicar ida y vuelta (×2)</span>
              </label>

              {previewDistance !== null && (
                <motion.div
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  className="bg-blue-50 border border-blue-100 rounded-lg p-3 space-y-1.5"
                >
                  <div className="flex justify-between text-xs">
                    <span className="text-slate-500">Distancia ida:</span>
                    <span className="font-semibold text-slate-700">{previewDistance.toFixed(2)} km</span>
                  </div>
                  <div className="flex justify-between text-xs">
                    <span className="text-slate-500">Trayecto ({isRoundTrip ? 'ida y vuelta' : 'ida'}):</span>
                    <span className="font-semibold text-slate-700">
                      {(previewDistance * (isRoundTrip ? 2 : 1)).toFixed(2)} km
                    </span>
                  </div>
                  <div className="flex justify-between text-xs border-t border-blue-100 pt-1.5 mt-1">
                    <span className="text-slate-500">Total estimado ({days} días):</span>
                    <span className="font-bold text-blue-700">
                      {(previewDistance * (isRoundTrip ? 2 : 1) * days).toFixed(2)} km
                    </span>
                  </div>
                </motion.div>
              )}
            </div>

            <div className="flex gap-2">
              <button
                onClick={handleCalculate}
                disabled={isCalculating || isSaving}
                className="flex-1 bg-white border border-slate-200 text-slate-700 rounded-lg py-3 text-sm font-semibold hover:bg-slate-50 transition-all active:scale-[0.98] disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {isCalculating ? <LoadingSpinner className="w-4 h-4" /> : <Calculator className="w-4 h-4" />}
                {isCalculating ? 'Calculando...' : 'Calcular'}
              </button>
              <button
                onClick={handleSave}
                disabled={isSaving || isCalculating}
                className="flex-1 bg-slate-900 text-white rounded-lg py-3 text-sm font-semibold hover:bg-slate-800 transition-all active:scale-[0.98] disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {isSaving ? <LoadingSpinner className="w-4 h-4" /> : editingId ? <CheckCircle2 className="w-4 h-4" /> : <Plus className="w-4 h-4" />}
                {isSaving ? 'Guardando...' : editingId ? 'Actualizar' : 'Guardar'}
              </button>
            </div>

            <AnimatePresence>
              {error && <ErrorBanner message={error} onClose={() => setError(null)} />}
              {success && <SuccessBanner message={success} />}
            </AnimatePresence>
          </section>

          {/* OneDrive Sync Card */}
          <section className="bg-slate-900 text-white rounded-2xl p-5 flex flex-col gap-3.5 shadow-md">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Cloud size={16} className="text-blue-400 shrink-0" />
                <h3 className="text-[11px] font-bold uppercase tracking-widest text-slate-300">Sincronización OneDrive</h3>
              </div>
              <div className={cn('w-2 h-2 rounded-full', onedriveStatus.connected ? 'bg-green-400 animate-pulse' : 'bg-amber-400')}></div>
            </div>

            {!onedriveStatus.isConfigured ? (
              <div className="space-y-2.5">
                <p className="text-[11px] text-slate-300 leading-relaxed">
                  Para habilitar la sincronización directa mensual en OneDrive, configura las variables de entorno.
                </p>
                <div className="bg-slate-850 p-2.5 rounded-lg space-y-1.5 text-[9px] text-slate-400 border border-slate-700/60">
                  <p className="font-bold text-slate-300">Instrucciones de configuración:</p>
                  <p>1. Crea un Registro de App en Azure (Microsoft Entra).</p>
                  <p>2. Añade la URI de redirección:</p>
                  <code className="block bg-slate-950 p-1.5 rounded text-[8px] text-blue-300 overflow-x-auto select-all break-all leading-normal">
                    {window.location.origin}/api/auth/microsoft/callback
                  </code>
                  <p>
                    3. Agrega <code>MICROSOFT_CLIENT_ID</code> y <code>MICROSOFT_CLIENT_SECRET</code> en el archivo <code>.env</code>.
                  </p>
                </div>
              </div>
            ) : !onedriveStatus.connected ? (
              <div className="space-y-3">
                <p className="text-[11px] text-slate-300 leading-relaxed">
                  Se generará automáticamente un archivo Excel mensual único por cada kilometraje ingresado por los consultores.
                </p>
                <button
                  onClick={handleConnectOneDrive}
                  className="w-full bg-blue-600 hover:bg-blue-500 text-white font-semibold py-2 px-3 rounded-lg text-xs flex items-center justify-center gap-2 transition-all cursor-pointer shadow-sm hover:scale-[1.01] active:scale-[0.99]"
                >
                  <Cloud className="w-4 h-4" />
                  Vincular OneDrive
                </button>
              </div>
            ) : (
              <div className="space-y-2.5">
                <div className="bg-slate-800/80 border border-slate-700/50 rounded-lg p-2.5">
                  <p className="text-[9px] uppercase font-bold tracking-wider text-slate-400">Cuenta Activa</p>
                  <p className="text-xs font-bold text-white truncate mt-0.5">{onedriveStatus.userName}</p>
                  <p className="text-[10px] text-slate-400 truncate">{onedriveStatus.userEmail}</p>
                </div>
                <p className="text-[10px] text-slate-400 italic leading-snug">
                  Guardando en: <code className="text-blue-300 bg-slate-950 px-1 py-0.5 rounded text-[8px]">/GeoConsult_Kilometraje/</code>
                </p>

                <div className="flex gap-2">
                  <button
                    onClick={handleBulkSync}
                    disabled={isSyncing}
                    className="flex-1 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 font-semibold py-1.5 px-2.5 rounded-lg text-xs flex items-center justify-center gap-1.5 transition-all cursor-pointer disabled:opacity-50"
                  >
                    <RefreshCw className={cn('w-3 h-3', isSyncing && 'animate-spin')} />
                    {isSyncing ? 'Sincronizando...' : 'Forzar Sinc.'}
                  </button>
                  <button
                    onClick={handleLogoutOneDrive}
                    className="bg-transparent hover:bg-red-950/30 text-red-400 border border-red-900/30 font-semibold py-1.5 px-2 rounded-lg text-xs flex items-center justify-center transition-all cursor-pointer"
                  >
                    Salir
                  </button>
                </div>
              </div>
            )}
          </section>

          {/* Excel Export Card */}
          <section className="bg-blue-50 rounded-2xl p-6 border border-blue-100 flex flex-col shadow-sm">
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-xs font-bold text-blue-600 uppercase tracking-widest">Exportación Local Excel</h3>
              <div className={cn('w-2 h-2 rounded-full', records.length > 0 ? 'bg-green-500 animate-pulse' : 'bg-slate-300')}></div>
            </div>
            <p className="text-[11px] text-blue-800 leading-relaxed mb-4">
              Descarga un reporte detallado y agrupado por consultor en tu ordenador local.
            </p>
            <button
              onClick={exportToExcel}
              disabled={filteredRecords.length === 0}
              className="bg-white rounded-lg p-3 border border-blue-200 flex items-center gap-3 w-full text-left hover:border-blue-300 transition-colors disabled:opacity-50 cursor-pointer shadow-sm"
            >
              <div className="w-8 h-8 bg-green-100 rounded flex items-center justify-center shrink-0">
                <FileSpreadsheet className="w-5 h-5 text-green-600" />
              </div>
              <div>
                <p className="text-[10px] font-bold text-slate-900 uppercase">Descargar Local</p>
                <p className="text-[9px] text-slate-400">Total: {filteredRecords.length} registros</p>
              </div>
            </button>
          </section>
        </div>

        {/* Main Content Area */}
        <div className="flex-1 flex flex-col gap-6 min-w-0 min-h-0">
          {/* Summary Cards */}
          <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 shrink-0">
            <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">Total Km Reportados</p>
              <p className="text-3xl font-black text-slate-900">
                {totalKm.toLocaleString()} <span className="text-sm font-medium text-slate-400">km</span>
              </p>
            </div>
            <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">Registros</p>
              <p className="text-3xl font-black text-slate-900">{filteredRecords.length}</p>
            </div>
            <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-1">Google Maps API</p>
              <div className="flex items-center gap-2">
                <span className="w-2 h-2 bg-green-500 rounded-full"></span>
                <span className="text-sm font-bold text-slate-700">Activa</span>
              </div>
            </div>
          </section>

          {/* Filters */}
          <section className="bg-white rounded-2xl border border-slate-200 p-4 shadow-sm shrink-0">
            <div className="flex flex-col sm:flex-row gap-3">
              <div className="relative flex-1">
                <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Filtrar por consultor..."
                  value={filters.consultant}
                  onChange={(e) => setFilters((f) => ({ ...f, consultant: e.target.value }))}
                  className="w-full border border-slate-200 rounded-lg pl-9 pr-3 py-2 text-sm placeholder:text-slate-300 outline-none focus:ring-2 focus:ring-blue-100"
                />
              </div>
              <div className="flex gap-3">
                <input
                  type="date"
                  value={filters.dateFrom}
                  onChange={(e) => setFilters((f) => ({ ...f, dateFrom: e.target.value }))}
                  className="border border-slate-200 rounded-lg px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-blue-100"
                />
                <input
                  type="date"
                  value={filters.dateTo}
                  onChange={(e) => setFilters((f) => ({ ...f, dateTo: e.target.value }))}
                  className="border border-slate-200 rounded-lg px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-blue-100"
                />
              </div>
            </div>
          </section>

          {/* History Table */}
          <section className="flex-1 bg-white rounded-2xl border border-slate-200 overflow-hidden shadow-sm flex flex-col min-h-0">
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50 shrink-0">
              <h2 className="text-xs font-bold uppercase tracking-widest text-slate-400">Historial de Reportes</h2>
              <span className="text-[10px] bg-slate-200 text-slate-600 px-2.5 py-1 rounded-full font-bold uppercase">
                {filteredRecords.length} Entradas
              </span>
            </div>

            <div className="flex-1 overflow-auto">
              {isLoadingRecords ? (
                <div className="flex flex-col items-center justify-center h-full gap-3 text-slate-400">
                  <LoadingSpinner className="w-8 h-8" />
                  <p className="text-xs">Cargando registros...</p>
                </div>
              ) : (
                <table className="w-full text-left border-collapse table-fixed">
                  <thead className="sticky top-0 bg-white border-b border-slate-100 z-20">
                    <tr className="text-[10px] uppercase tracking-wider text-slate-400">
                      <th className="px-6 py-4 font-bold w-[18%]">Consultor</th>
                      <th className="px-6 py-4 font-bold w-[35%] text-center">Ruta</th>
                      <th className="px-6 py-4 font-bold w-[10%] text-center">Días</th>
                      <th className="px-6 py-4 font-bold w-[12%] text-center">Fecha</th>
                      <th className="px-6 py-4 font-bold w-[15%] text-right">Km Totales</th>
                      <th className="px-6 py-4 w-[15%]"></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50 text-sm">
                    <AnimatePresence>
                      {filteredRecords.map((record) => (
                        <motion.tr
                          layout
                          key={record.id}
                          initial={{ opacity: 0 }}
                          animate={{ opacity: 1 }}
                          className="hover:bg-slate-50/50 transition-colors group"
                        >
                          <td className="px-6 py-4 font-semibold text-slate-900 truncate">{record.consultantName}</td>
                          <td className="px-6 py-4">
                            <div className="flex items-center gap-2 justify-center">
                              <span className="text-[11px] text-slate-400 truncate max-w-[120px]">{record.origin}</span>
                              <ArrowRightLeft size={12} className="text-slate-200 shrink-0" />
                              <span className="text-[11px] text-slate-600 font-medium truncate max-w-[120px]">{record.destination}</span>
                            </div>
                          </td>
                          <td className="px-6 py-4 text-center text-slate-500 font-medium">{record.days}</td>
                          <td className="px-6 py-4 text-center text-slate-500 text-[11px]">{record.date}</td>
                          <td className="px-6 py-4 text-right font-mono font-bold text-blue-600">
                            <div className="flex items-center justify-end gap-2">
                              <span>{record.totalDistanceKm.toFixed(2)}</span>
                              {record.syncedToOneDrive ? (
                                <span
                                  title="Sincronizado en tiempo real con OneDrive"
                                  className="text-green-500 hover:scale-110 transition-all cursor-help shrink-0"
                                >
                                  <Cloud size={14} className="inline" />
                                </span>
                              ) : (
                                <span
                                  title="Guardado localmente. Vincula OneDrive para sincronizar en la nube."
                                  className="text-slate-300 hover:scale-110 transition-all cursor-help shrink-0"
                                >
                                  <CloudOff size={14} className="inline" />
                                </span>
                              )}
                            </div>
                          </td>
                          <td className="px-6 py-4 text-right">
                            <div className="flex items-center justify-end gap-1 opacity-0 group-hover:opacity-100 transition-opacity">
                              <button
                                onClick={() => handleEdit(record)}
                                className="p-1.5 text-slate-300 hover:text-blue-600 transition-colors"
                                title="Editar"
                              >
                                <Edit2 size={14} />
                              </button>
                              <button
                                onClick={() => handleDelete(record.id)}
                                className="p-1.5 text-slate-300 hover:text-red-500 transition-colors"
                                title="Eliminar"
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>
                          </td>
                        </motion.tr>
                      ))}
                    </AnimatePresence>
                    {filteredRecords.length === 0 && (
                      <tr>
                        <td colSpan={6} className="px-6 py-12 text-center text-slate-300 text-xs italic">
                          No hay registros disponibles para mostrar.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              )}
            </div>
          </section>
        </div>
      </main>
    </div>
  );
}
