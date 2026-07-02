import React, { useState, useCallback, useRef, useEffect } from 'react';
import { Toaster, toast } from 'sonner';
import {
  TrendingUp, Play, Download, Sparkles, Upload, FileSpreadsheet, User, X
} from 'lucide-react';
import { FileUpload } from './components/FileUpload';
import { ParameterForm } from './components/ParameterForm';
import { SimulationProgress } from './components/SimulationProgress';
import { CorpusSurvivalChart } from './components/CorpusSurvivalChart';
import { ResultsSummary } from './components/ResultsSummary';
import { ExpenseProfileManager } from './components/ExpenseProfileManager';
import {
  parseQuarterlyReturnsExcel,
  parseUserExpensesExcel,
  generateResultsExcel,
  parseResultsExcel,
  generateSampleExpenseExcel,
} from './engine/excelParser';
import { calculateChartData, type ChartData } from './engine/chartCalculations';
import type { AssetReturns, SimulationParams, YearlySummaryData } from './engine/simulation';
import type { UserExpenseRow } from './engine/excelParser';
import type { WorkerResponse } from './engine/simulationWorker';
import {
  fetchQuarterlyReturns,
  saveQuarterlyReturns,
  deleteQuarterlyReturns,
  fetchParameters,
  saveParameters,
  fetchExpenseProfiles,
  saveExpenseProfile,
  deleteExpenseProfile,
  fetchSavedSimulation,
  saveSimulation,
} from './lib/api';

function App() {
  const [activeUserId, setActiveUserId] = useState<string | null>(null);
  const [userIdInput, setUserIdInput] = useState('');

  // ── Data state ──
  const [assetReturns, setAssetReturns] = useState<AssetReturns | null>(null);
  const [quarterlyFileName, setQuarterlyFileName] = useState<string>('');
  const [quarterlyError, setQuarterlyError] = useState<string | null>(null);
  const [viewingQuarterly, setViewingQuarterly] = useState(false);

  const [expenseProfiles, setExpenseProfiles] = useState<Array<{ name: string; data: UserExpenseRow[]; dict: Record<number, number> }>>([]);
  const [selectedProfile, setSelectedProfile] = useState<string | null>(null);
  const [expenseError, setExpenseError] = useState<string | null>(null);

  // ── localStorage helpers ──
  const lsGet = (key: string) => { try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : null; } catch { return null; } };
  const lsSet = (key: string, val: any) => { try { localStorage.setItem(key, JSON.stringify(val)); } catch { console.warn('localStorage write failed (quota exceeded?)'); } };
  const lsDel = (key: string) => { try { localStorage.removeItem(key); } catch {} };

  // ── Load parameters, quarterly returns & expenses when user changes ──
  useEffect(() => {
    if (!activeUserId) return;
    let cancelled = false;

    // Clear the previous session's data immediately so a user switch never shows stale leftovers while fetching
    setAssetReturns(null);
    setQuarterlyFileName('');

    const applyParams = (p: any) => {
      setAge(p.age);
      setWithdrawalStartAge(p.withdrawal_start_age);
      setWithdrawalAmount(p.withdrawal_amount);
      setWithdrawalYear(p.withdrawal_year);
      setInitialCorpus(p.initial_corpus);
      setEquityAllocation(p.equity_allocation);
      setRealEstateAllocation(p.real_estate_allocation);
      setPassiveAllocation(p.passive_allocation);
      setDebtAllocation(p.debt_allocation);
      setAltAllocation(p.alt_allocation);
    };
    const resetParams = () => {
      setAge('');
      setWithdrawalStartAge('');
      setInitialCorpus('');
      setEquityAllocation('');
      setRealEstateAllocation('');
      setPassiveAllocation('');
      setDebtAllocation('');
      setAltAllocation('');
    };

    (async () => {
      try {
        const res = await fetchParameters(activeUserId);
        if (cancelled) return;
        if (res && res.age !== undefined && res.age !== null) {
          applyParams(res);
          lsSet(`mc_params_${activeUserId}`, { user_id: activeUserId, ...res });
        } else {
          resetParams();
          lsDel(`mc_params_${activeUserId}`);
        }
      } catch {
        const saved = lsGet(`mc_params_${activeUserId}`);
        if (cancelled) return;
        if (saved) {
          applyParams(saved);
          toast.warning('Using locally cached parameters — server unreachable');
        } else {
          resetParams();
        }
      }
    })();

    const applyQuarterly = (data: any[], fileName: string) => {
      const ar = {
        equity: data.map((r: any) => r.equity),
        realEstate: data.map((r: any) => r.realEstate),
        commodity: data.map((r: any) => r.commodity),
        debt: data.map((r: any) => r.debt),
        alternative: data.map((r: any) => r.alternative),
      };
      setAssetReturns(ar);
      setQuarterlyFileName(fileName);
    };

    (async () => {
      try {
        const res = await fetchQuarterlyReturns();
        if (cancelled) return;
        if (res.data?.length > 0) {
          applyQuarterly(res.data, res.fileName || '');
          lsSet('mc_quarterly_global', { data: res.data, fileName: res.fileName });
        } else {
          setAssetReturns(null);
          setQuarterlyFileName('');
          lsDel('mc_quarterly_global');
        }
      } catch {
        const saved = lsGet('mc_quarterly_global');
        if (cancelled) return;
        if (saved?.data?.length > 0) {
          applyQuarterly(saved.data, saved.fileName);
          toast.warning('Using locally cached quarterly returns — server unreachable');
        } else {
          setAssetReturns(null);
          setQuarterlyFileName('');
        }
      }
    })();

    (async () => {
      try {
        const profiles = await fetchExpenseProfiles(activeUserId);
        if (cancelled) return;
        if (profiles.length > 0) {
          setExpenseProfiles(profiles);
          setSelectedProfile(profiles[0].name);
          lsSet(`mc_expenses_${activeUserId}`, profiles);
        } else {
          setExpenseProfiles([]);
          setSelectedProfile(null);
        }
      } catch {
        const cached = lsGet(`mc_expenses_${activeUserId}`) || [];
        if (cancelled) return;
        if (cached.length > 0) {
          setExpenseProfiles(cached);
          setSelectedProfile(cached[0].name);
          toast.warning('Using locally cached expense profiles — server unreachable');
        } else {
          setExpenseProfiles([]);
          setSelectedProfile(null);
        }
      }
    })();

    return () => { cancelled = true; };
  }, [activeUserId]);



  // ── Parameter state ──
  const [age, setAge] = useState<number | ''>(42);
  const [withdrawalStartAge, setWithdrawalStartAge] = useState<number | ''>(42);
  const [withdrawalAmount, setWithdrawalAmount] = useState<number | ''>(3);
  const [withdrawalYear, setWithdrawalYear] = useState<number | ''>(2034);
  const [initialCorpus, setInitialCorpus] = useState<number | ''>(58994959);
  const [equityAllocation, setEquityAllocation] = useState<number | ''>(0.52);
  const [realEstateAllocation, setRealEstateAllocation] = useState<number | ''>(0.14);
  const [passiveAllocation, setPassiveAllocation] = useState<number | ''>(0);
  const [debtAllocation, setDebtAllocation] = useState<number | ''>(0.23);
  const [altAllocation, setAltAllocation] = useState<number | ''>(0.11);
  const [simulationType, setSimulationType] = useState<'expenses' | 'swr'>('expenses');
  const [activeSWR, setActiveSWR] = useState<number>(0.04);
  const [swrInitial, setSwrInitial] = useState<number | ''>(4);
  const [swrStepUp, setSwrStepUp] = useState<number | ''>(0.5);
  const [swrFinal, setSwrFinal] = useState<number | ''>(7);

  const baseSWRList: number[] = (() => {
    const initial = Math.max(Number(swrInitial) || 4, 0.1);
    const step = Math.max(Number(swrStepUp) || 0.5, 0.1);
    const final = Math.max(Number(swrFinal) || 7, initial);
    const list: number[] = [];
    let pct = Math.round(initial * 100) / 100;
    const finalRounded = Math.round(final * 100) / 100;
    const stepRounded = Math.round(step * 100) / 100;
    while (pct <= finalRounded + 0.001) {
      list.push(parseFloat((pct / 100).toFixed(6)));
      pct = Math.round((pct + stepRounded) * 100) / 100;
    }
    return list.length > 0 ? list : [0.04];
  })();

  // ── Sync activeSWR to swrInitial when it changes ──
  useEffect(() => {
    const initial = Math.max(Number(swrInitial) || 4, 0.1);
    setActiveSWR(parseFloat((initial / 100).toFixed(6)));
  }, [swrInitial]);

  // ── Auto-calculate Withdrawal Year ──
  useEffect(() => {
    const currentYear = new Date().getFullYear();
    setWithdrawalYear(currentYear + (Number(withdrawalStartAge) || 0) - (Number(age) || 0));
  }, [age, withdrawalStartAge]);

  // ── Auto-calculate Withdrawal Amount ──
  useEffect(() => {
    if (simulationType === 'expenses') {
      const profile = expenseProfiles.find(p => p.name === selectedProfile);
      const startAgeNum = Number(withdrawalStartAge) || 0;
      if (profile && profile.dict[startAgeNum] !== undefined) {
        const annualRaw = profile.dict[startAgeNum];
        const monthlyLakhs = (annualRaw / 12) / 100000;
        setWithdrawalAmount(monthlyLakhs);
      }
    } else {
      const annualRaw = (Number(initialCorpus) || 0) * activeSWR;
      const monthlyLakhs = (annualRaw / 12) / 100000;
      setWithdrawalAmount(monthlyLakhs);
    }
  }, [simulationType, expenseProfiles, selectedProfile, withdrawalStartAge, initialCorpus, activeSWR]);

  // ── Auto-save parameters to server (per user ID), with localStorage as an offline cache ──
  useEffect(() => {
    if (!activeUserId) return;
    const timeoutId = setTimeout(() => {
      const params = {
        age: Number(age) || 0,
        withdrawal_start_age: Number(withdrawalStartAge) || 0,
        withdrawal_amount: Number(withdrawalAmount) || 0,
        withdrawal_year: Number(withdrawalYear) || 0,
        initial_corpus: Number(initialCorpus) || 0,
        equity_allocation: Number(equityAllocation) || 0,
        real_estate_allocation: Number(realEstateAllocation) || 0,
        passive_allocation: Number(passiveAllocation) || 0,
        debt_allocation: Number(debtAllocation) || 0,
        alt_allocation: Number(altAllocation) || 0,
      };
      lsSet(`mc_params_${activeUserId}`, { user_id: activeUserId, ...params });
      saveParameters(activeUserId, params).catch(() => {
        toast.warning('Saved locally, but failed to sync parameters to server');
      });
    }, 1000);
    return () => clearTimeout(timeoutId);
  }, [
    activeUserId, age, withdrawalStartAge, withdrawalAmount, withdrawalYear, initialCorpus,
    equityAllocation, realEstateAllocation, passiveAllocation, debtAllocation, altAllocation
  ]);

  // ── Simulation state ──
  const [isSimulating, setIsSimulating] = useState(false);
  const [progress, setProgress] = useState<{ current: number; total: number; percentage: number } | null>(null);
  const [results, setResults] = useState<Array<{ baseSWR: number; survivalRate: number; exhaustionProbability: number; numFailures: number }>>([]);
  const [chartData, setChartData] = useState<ChartData | null>(null);
  const [yearlySummary, setYearlySummary] = useState<YearlySummaryData[]>([]);
  const workerRef = useRef<Worker | null>(null);
  const [chartBaseRank, setChartBaseRank] = useState<number>(7500);
  const [excelYearlySummary, setExcelYearlySummary] = useState<YearlySummaryData[] | null>(null);
  const [savedType2Charts, setSavedType2Charts] = useState<Record<number, ChartData> | null>(null);

  // ── Load saved chart when profile or simulation type changes ──
  useEffect(() => {
    const effectiveProfile = simulationType === 'swr' ? '__TYPE_2_SWR__' : selectedProfile;

    if (!activeUserId || !effectiveProfile) {
      setChartData(null);
      return;
    }

    let cancelled = false;

    // Clear any currently loaded simulation data since the profile changed
    setYearlySummary([]);
    setExcelYearlySummary(null);
    setChartFromExcel(null);
    setChartData(null);
    setResults([]);
    setSavedType2Charts(null);

    const applyData = (data: any) => {
      if (simulationType === 'swr') {
        setSavedType2Charts(data.chartData);
        setYearlySummary([]);
      } else {
        setChartData(data.chartData);
        setYearlySummary(data.yearlySummary || []);
        setSavedType2Charts(null);
      }
      setResults(data.swrResults);
    };

    (async () => {
      try {
        const result = await fetchSavedSimulation(activeUserId, effectiveProfile);
        if (cancelled) return;
        if (result?.chartData && result?.swrResults) {
          applyData(result);
          lsSet(`mc_sim_${activeUserId}_${effectiveProfile}`, result);
        } else {
          setChartData(null);
          setResults([]);
          setSavedType2Charts(null);
        }
      } catch {
        const cached = lsGet(`mc_sim_${activeUserId}_${effectiveProfile}`);
        if (cancelled) return;
        if (cached?.chartData && cached?.swrResults) {
          applyData(cached);
          toast.warning('Using locally cached chart — server unreachable');
        } else {
          setChartData(null);
          setResults([]);
          setSavedType2Charts(null);
        }
      }
    })();

    return () => { cancelled = true; };
  }, [activeUserId, selectedProfile, simulationType]);

  // ── Auto-derive chart data when yearlySummary, activeSWR, rank, or type changes ──
  useEffect(() => {
    if (yearlySummary && yearlySummary.length > 0) {
      const targetSWR = simulationType === 'expenses' ? 0.04 : activeSWR;
      const filteredSummary = yearlySummary.filter(d => d.baseSWR === targetSWR);
      if (filteredSummary.length > 0) {
        setChartData(calculateChartData(filteredSummary, chartBaseRank));
      } else {
        setChartData(null);
      }
    } else if (simulationType === 'swr' && savedType2Charts) {
      setChartData(savedType2Charts[activeSWR] || null);
    }
  }, [yearlySummary, chartBaseRank, activeSWR, simulationType, savedType2Charts]);

  // ── Upload for chart from existing Excel ──
  const [chartFromExcel, setChartFromExcel] = useState<ChartData | null>(null);

  const totalAllocation = equityAllocation + realEstateAllocation + passiveAllocation + debtAllocation + altAllocation;
  const allocationValid = Math.abs(totalAllocation - 1.0) < 0.0001;

  // ── Handlers ──
  const handleQuarterlyFile = useCallback((buffer: ArrayBuffer, fileName: string) => {
    try {
      setQuarterlyError(null);
      const { assetReturns: ar, data } = parseQuarterlyReturnsExcel(buffer);
      setAssetReturns(ar);
      setQuarterlyFileName(fileName);
      lsSet('mc_quarterly_global', { data, fileName });
      saveQuarterlyReturns(data, fileName).catch(() => {
        toast.warning('Saved locally, but failed to sync quarterly returns to server');
      });
      toast.success(`Loaded ${ar.equity.length} quarterly return records`);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to parse file';
      setQuarterlyError(msg);
      toast.error(msg);
    }
  }, []);

  const handleExpenseFile = useCallback((buffer: ArrayBuffer, fileName: string) => {
    try {
      setExpenseError(null);
      const { data, expensesDict } = parseUserExpensesExcel(buffer);
      const newProfile = { name: fileName, data, dict: expensesDict };
      setExpenseProfiles(prev => {
        const updated = [...prev.filter(p => p.name !== fileName), newProfile];
        if (activeUserId) lsSet(`mc_expenses_${activeUserId}`, updated);
        return updated;
      });
      setSelectedProfile(fileName);
      if (activeUserId) {
        saveExpenseProfile(activeUserId, fileName, fileName, data).catch(() => {
          toast.warning('Saved locally, but failed to sync expense profile to server');
        });
      }
      toast.success(`Loaded ${data.length} expense entries from ${fileName}`);
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to parse file';
      setExpenseError(msg);
      toast.error(msg);
    }
  }, [activeUserId]);

  const handleDeleteProfile = useCallback((name: string) => {
    setExpenseProfiles(prev => {
      const updated = prev.filter(p => p.name !== name);
      if (activeUserId) lsSet(`mc_expenses_${activeUserId}`, updated);
      return updated;
    });
    if (selectedProfile === name) setSelectedProfile(null);
    if (activeUserId) {
      deleteExpenseProfile(activeUserId, name).catch(() => {
        toast.warning('Deleted locally, but failed to delete from server');
      });
    }
    toast.success('Profile deleted');
  }, [selectedProfile, activeUserId]);

  const handleRunSimulation = useCallback(() => {
    if (!assetReturns) {
      toast.error('Please upload quarterly returns data first');
      return;
    }
    let profileDict: Record<number, number> = {};
    if (simulationType === 'expenses') {
      const profile = expenseProfiles.find(p => p.name === selectedProfile);
      if (!profile) {
        toast.error('Please select an expense profile');
        return;
      }
      profileDict = profile.dict;
    }
    if (!allocationValid) {
      toast.error('Asset allocations must sum to 100%');
      return;
    }

    // Save parameters to server (per user ID), with localStorage as an offline cache
    if (activeUserId) {
      const params = {
        age: Number(age) || 0,
        withdrawal_start_age: Number(withdrawalStartAge) || 0,
        withdrawal_amount: Number(withdrawalAmount) || 0,
        withdrawal_year: Number(withdrawalYear) || 0,
        initial_corpus: Number(initialCorpus) || 0,
        equity_allocation: Number(equityAllocation) || 0,
        real_estate_allocation: Number(realEstateAllocation) || 0,
        passive_allocation: Number(passiveAllocation) || 0,
        debt_allocation: Number(debtAllocation) || 0,
        alt_allocation: Number(altAllocation) || 0,
      };
      lsSet(`mc_params_${activeUserId}`, { user_id: activeUserId, ...params });
      saveParameters(activeUserId, params).catch(() => {
        toast.warning('Saved locally, but failed to sync parameters to server');
      });
    }

    // Reset previous results
    setResults([]);
    setChartData(null);
    setYearlySummary([]);
    setChartFromExcel(null);
    setIsSimulating(true);
    setProgress({ current: 0, total: 10000, percentage: 0 });

    const params: SimulationParams = {
      initialCorpus: Number(initialCorpus) || 0,
      currentAge: Number(age) || 0,
      withdrawalStartAge: Number(withdrawalStartAge) || 0,
      targetAge: 91,
      equityAllocation: Number(equityAllocation) || 0,
      realEstateAllocation: Number(realEstateAllocation) || 0,
      commodityAllocation: Number(passiveAllocation) || 0,
      debtAllocation: Number(debtAllocation) || 0,
      alternativeAllocation: Number(altAllocation) || 0,
      baseSWRList,
      swrInflationRate: 0.06,
      numSimulations: 10000,
      simulationType,
      expensesDict: profileDict,
      assetReturns,
    };

    // Create Web Worker
    const worker = new Worker(
      new URL('./engine/simulationWorker.ts', import.meta.url),
      { type: 'module' }
    );
    workerRef.current = worker;

    worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
      const msg = event.data;

      if (msg.type === 'progress') {
        setProgress({ current: msg.current, total: msg.total, percentage: msg.percentage });
      } else if (msg.type === 'complete') {
        setIsSimulating(false);
        setProgress(null);
        worker.terminate();
        workerRef.current = null;

        const output = msg.results;
        const summaryResults = output.results.map(r => ({
          baseSWR: r.baseSWR,
          survivalRate: r.survivalRate,
          exhaustionProbability: r.exhaustionProbability,
          numFailures: r.numFailures
        }));
        setResults(summaryResults);

        // Combine yearly summaries for all SWRs
        const allYearlySummary = output.results.flatMap(r => r.yearlySummary);
        
        if (allYearlySummary.length > 0) {
          setExcelYearlySummary(null);
          setYearlySummary(allYearlySummary);

          // Calculate chart data for the first SWR in the list to satisfy the API
          const firstSWR = baseSWRList[0] || 0.04;
          const filtered04 = allYearlySummary.filter(d => d.baseSWR === firstSWR);
          const cd = filtered04.length > 0 ? calculateChartData(filtered04, chartBaseRank) : null;
          
          let payloadChartData: any = cd;
          let payloadYearlySummary: any[] = allYearlySummary;
          
          if (simulationType === 'swr') {
            payloadYearlySummary = []; // Prevent 250MB stringify crash!
            const type2Charts: Record<number, ChartData> = {};
            for (const swr of baseSWRList) {
              const swrData = allYearlySummary.filter(d => d.baseSWR === swr);
              if (swrData.length > 0) {
                type2Charts[swr] = calculateChartData(swrData, chartBaseRank);
              }
            }
            payloadChartData = type2Charts;
            setSavedType2Charts(type2Charts);
            // Default activeSWR chart on screen right after running
            setChartData(type2Charts[activeSWR] || null);
          } else {
            setSavedType2Charts(null);
            setChartData(cd);
          }

          const effectiveProfile = simulationType === 'swr' ? '__TYPE_2_SWR__' : selectedProfile;
          if (activeUserId && effectiveProfile) {
            // yearlySummary (~34MB) stays in-memory for rank adjustment this session
            // but is never persisted — too large for both localStorage and the server
            const simPayload = {
              chartData: payloadChartData,
              yearlySummary: [],
              swrResults: summaryResults,
            };
            lsSet(`mc_sim_${activeUserId}_${effectiveProfile}`, simPayload);
            saveSimulation(activeUserId, effectiveProfile, simPayload).catch(() => {
              toast.warning('Chart saved locally, but failed to sync to server');
            });
          }

          toast.success('Simulation complete!');
        } else {
          toast.success('Simulation complete!');
        }
      } else if (msg.type === 'error') {
        setIsSimulating(false);
        setProgress(null);
        worker.terminate();
        workerRef.current = null;
        toast.error(msg.message);
      }
    };

    worker.onerror = (error) => {
      setIsSimulating(false);
      setProgress(null);
      worker.terminate();
      workerRef.current = null;
      toast.error('Simulation worker error: ' + error.message);
    };

    worker.postMessage({ type: 'start', params });
  }, [
    assetReturns, expenseProfiles, selectedProfile, allocationValid,
    initialCorpus, age, withdrawalStartAge, withdrawalAmount, withdrawalYear,
    equityAllocation, realEstateAllocation, passiveAllocation, debtAllocation, altAllocation,
    simulationType, activeUserId, activeSWR, chartBaseRank
  ]);

  const handleDownloadExcel = useCallback(() => {
    if (yearlySummary.length === 0) return;
    const blob = generateResultsExcel(yearlySummary);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `simulation_results_${new Date().toISOString().split('T')[0]}.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    toast.success('Excel file downloaded');
  }, [yearlySummary]);

  const handleDownloadSampleExpense = useCallback(() => {
    const blob = generateSampleExpenseExcel();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `Sample_Expense_Profile.xlsx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    toast.success('Sample format downloaded');
  }, []);

  const handleChartFromExcel = useCallback((buffer: ArrayBuffer, fileName: string) => {
    try {
      const data = parseResultsExcel(buffer);
      setExcelYearlySummary(data);

      const uniqueSWRs = [...new Set(data.map(d => d.baseSWR))].sort((a, b) => a - b);
      if (uniqueSWRs.length === 0) {
        toast.error('No data found in Excel file');
        return;
      }

      const targetSWR = uniqueSWRs.includes(activeSWR) ? activeSWR : uniqueSWRs[0];
      const filteredSummary = data.filter(d => d.baseSWR === targetSWR);

      if (filteredSummary.length > 0) {
        setChartFromExcel(calculateChartData(filteredSummary, chartBaseRank));
        toast.success('Chart generated from Excel file');
      } else {
        toast.error('No matching data found in Excel file');
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : 'Failed to parse Excel';
      toast.error(msg);
    }
  }, [activeSWR, chartBaseRank]);

  // ── Validation ──
  const hasData = !!assetReturns && (simulationType === 'swr' || (expenseProfiles.length > 0 && !!selectedProfile));
  const canRun = !!assetReturns && (simulationType === 'swr' || !!selectedProfile) && allocationValid && !isSimulating;
  const displayChart = chartFromExcel || chartData;

  if (!activeUserId) {
    return (
      <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--color-bg-base)' }}>
        <div className="glass-card" style={{ padding: '2.5rem', width: '100%', maxWidth: 400, textAlign: 'center' }}>
          <div style={{
            width: 56, height: 56, borderRadius: 'var(--radius-md)', margin: '0 auto 1.5rem',
            background: 'var(--gradient-primary)',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <TrendingUp size={32} color="white" />
          </div>
          <h2 style={{ fontSize: '1.5rem', fontWeight: 800, marginBottom: '0.5rem' }}>Welcome</h2>
          <p style={{ color: 'var(--color-text-muted)', marginBottom: '1.5rem', fontSize: '0.875rem' }}>
            Please enter your User ID to access your dashboard and saved simulations.
          </p>
          <form onSubmit={(e) => {
            e.preventDefault();
            if (userIdInput.trim()) {
              setActiveUserId(userIdInput.trim());
            }
          }}>
            <input
              type="text"
              className="input-field"
              placeholder="Enter User ID..."
              value={userIdInput}
              onChange={(e) => setUserIdInput(e.target.value)}
              style={{ marginBottom: '1rem', width: '100%', textAlign: 'center' }}
              autoFocus
            />
            <button type="submit" className="btn-primary" style={{ width: '100%' }} disabled={!userIdInput.trim()}>
              <User size={16} />
              Continue
            </button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: '100vh' }}>
      <Toaster
        position="top-right"
        toastOptions={{
          style: {
            background: 'var(--color-bg-card)',
            color: 'var(--color-text-primary)',
            border: '1px solid var(--color-border)',
          },
        }}
      />

      {/* Navigation */}
      <nav style={{
        position: 'sticky', top: 0, zIndex: 50,
        background: 'rgba(10,14,26,0.8)',
        backdropFilter: 'blur(20px)',
        borderBottom: '1px solid var(--color-border)',
      }}>
        <div className="container" style={{
          display: 'flex', justifyContent: 'space-between', alignItems: 'center',
          padding: '0.875rem 1.5rem',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
            <div style={{
              width: 36, height: 36, borderRadius: 'var(--radius-md)',
              background: 'var(--gradient-primary)',
              display: 'flex', alignItems: 'center', justifyContent: 'center',
            }}>
              <TrendingUp size={20} color="white" />
            </div>
            <div>
              <h1 style={{ fontSize: '1.125rem', fontWeight: 700, lineHeight: 1.2 }}>
                Retirement Planner
              </h1>
              <p style={{ fontSize: '0.6875rem', color: 'var(--color-text-muted)', letterSpacing: '0.05em', textTransform: 'uppercase' }}>
                Monte Carlo Simulation
              </p>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
            {yearlySummary.length > 0 && (
              <button className="btn-secondary" onClick={handleDownloadExcel}>
                <Download size={16} />
                Export Excel
              </button>
            )}
            <button className="btn-secondary" onClick={() => {
              setActiveUserId(null);
              setUserIdInput('');
              setResults([]);
              setChartData(null);
            }} title="Switch User">
              <User size={16} />
              <span style={{ maxWidth: 100, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {activeUserId}
              </span>
            </button>
          </div>
        </div>
      </nav>

      {/* Hero */}
      <div style={{
        textAlign: 'center', padding: '3rem 1.5rem 2rem',
        background: 'var(--gradient-glow)',
      }}>
        <div style={{ display: 'inline-flex', alignItems: 'center', gap: '0.5rem', marginBottom: '0.75rem' }}>
          <Sparkles size={16} color="var(--color-accent-amber)" />
          <span style={{ fontSize: '0.8125rem', fontWeight: 600, color: 'var(--color-accent-amber)', letterSpacing: '0.05em' }}>
            10,000 ITERATIONS • CLIENT-SIDE PROCESSING
          </span>
        </div>
        <h2 style={{ fontSize: '2rem', fontWeight: 800, marginBottom: '0.5rem' }}>
          Plan Your Retirement with{' '}
          <span style={{ background: 'var(--gradient-primary)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent' }}>
            Confidence
          </span>
        </h2>
        <p style={{ fontSize: '1rem', color: 'var(--color-text-secondary)', maxWidth: 600, margin: '0 auto' }}>
          Upload your data, configure parameters, and run Monte Carlo simulations to assess corpus survival probability.
        </p>
      </div>

      {/* Main Content */}
      <main className="container" style={{ paddingTop: '1.5rem', paddingBottom: '4rem' }}>
        <div className="stagger-children" style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem' }}>

          {/* Step 1: File Uploads */}
          <div className="glass-card" style={{ padding: '1.5rem' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', marginBottom: '1.25rem' }}>
              <div style={{
                width: 32, height: 32, borderRadius: '50%',
                background: 'var(--gradient-primary)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: '0.875rem', fontWeight: 700, color: 'white',
              }}>1</div>
              <h3>Upload Data</h3>
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: simulationType === 'swr' ? '1fr' : '1fr 1fr', gap: '1.5rem' }}>
              <FileUpload
                label="Quarterly Returns"
                description="Historical quarterly returns for 5 asset classes"
                onFileLoaded={handleQuarterlyFile}
                isLoaded={!!assetReturns}
                loadedFileName={quarterlyFileName}
                onView={() => setViewingQuarterly(true)}
                onClear={() => {
                  setAssetReturns(null);
                  setQuarterlyFileName('');
                  lsDel('mc_quarterly_global');
                  deleteQuarterlyReturns().catch(() => {
                    toast.warning('Cleared locally, but failed to delete from server');
                  });
                }}
                error={quarterlyError}
              />

            </div>
          </div>


          {/* Step 2: Parameters */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', marginBottom: '1rem' }}>
              <div style={{
                width: 32, height: 32, borderRadius: '50%',
                background: 'var(--gradient-primary)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: '0.875rem', fontWeight: 700, color: 'white',
              }}>2</div>
              <h3>Configure Parameters</h3>
            </div>
            <ParameterForm
              age={age} setAge={setAge}
              withdrawalStartAge={withdrawalStartAge} setWithdrawalStartAge={setWithdrawalStartAge}

              withdrawalYear={withdrawalYear} setWithdrawalYear={setWithdrawalYear}
              initialCorpus={initialCorpus} setInitialCorpus={setInitialCorpus}
              equityAllocation={equityAllocation} setEquityAllocation={setEquityAllocation}
              realEstateAllocation={realEstateAllocation} setRealEstateAllocation={setRealEstateAllocation}
              passiveAllocation={passiveAllocation} setPassiveAllocation={setPassiveAllocation}
              debtAllocation={debtAllocation} setDebtAllocation={setDebtAllocation}
              altAllocation={altAllocation} setAltAllocation={setAltAllocation}
              simulationType={simulationType} setSimulationType={setSimulationType}
            />
          </div>

          {/* Simulation Type Logic */}
          <div className="glass-card" style={{ padding: '1.5rem', marginTop: '-0.5rem' }}>
            <label className="input-label" style={{ marginBottom: '0.75rem', fontSize: '1rem', color: 'var(--color-text-primary)' }}>Simulation Type Logic</label>
            <select 
              className="input-field" 
              value={simulationType}
              onChange={(e) => setSimulationType(e.target.value as 'expenses' | 'swr')}
              style={{ width: '100%', cursor: 'pointer' }}
            >
              <option value="expenses">Type 1: User Expenses Based</option>
              <option value="swr">Type 2: Safe Withdrawal Rate (SWR) Based</option>
            </select>
          </div>

          {/* SWR Range Inputs (Type 2 only) */}
          {simulationType === 'swr' && (
            <div className="glass-card" style={{ padding: '1.5rem', marginTop: '-0.5rem' }}>
              <label className="input-label" style={{ marginBottom: '0.75rem', fontSize: '1rem', color: 'var(--color-text-primary)' }}>SWR Range</label>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: '1rem' }}>
                <div>
                  <label className="input-label">Initial SWR (%)</label>
                  <input
                    type="number"
                    className="input-field"
                    value={swrInitial}
                    onChange={e => setSwrInitial(e.target.value === '' ? '' : parseFloat(e.target.value))}
                    min={0.1}
                    max={30}
                    step={0.1}
                  />
                </div>
                <div>
                  <label className="input-label">Step Up (%)</label>
                  <input
                    type="number"
                    className="input-field"
                    value={swrStepUp}
                    onChange={e => setSwrStepUp(e.target.value === '' ? '' : parseFloat(e.target.value))}
                    min={0.1}
                    max={10}
                    step={0.1}
                  />
                </div>
                <div>
                  <label className="input-label">Final SWR (%)</label>
                  <input
                    type="number"
                    className="input-field"
                    value={swrFinal}
                    onChange={e => setSwrFinal(e.target.value === '' ? '' : parseFloat(e.target.value))}
                    min={0.1}
                    max={30}
                    step={0.1}
                  />
                </div>
              </div>
              <p style={{ margin: '0.75rem 0 0', fontSize: '0.8125rem', color: 'var(--color-text-muted)' }}>
                Will simulate: {baseSWRList.map(s => (s * 100).toFixed(1) + '%').join(', ')}
              </p>
            </div>
          )}

          {/* User Expenses Section */}
          {simulationType === 'expenses' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: '1.5rem', position: 'relative' }}>
              <div style={{ position: 'absolute', right: '1rem', top: '0.75rem', zIndex: 10 }}>
                <button 
                  className="btn-secondary" 
                  onClick={handleDownloadSampleExpense} 
                  style={{ padding: '0.25rem 0.5rem', fontSize: '0.75rem', height: '26px', minHeight: '26px', gap: '0.25rem' }}
                >
                  <Download size={12} /> Sample Format
                </button>
              </div>
              <FileUpload
                label="User Expenses"
                description="Age-based annual withdrawal amounts"
                onFileLoaded={handleExpenseFile}
                isLoaded={false}
                error={expenseError}
              />
              {expenseProfiles.length > 0 && (
                <ExpenseProfileManager
                  profiles={expenseProfiles}
                  selectedProfile={selectedProfile}
                  onSelect={setSelectedProfile}
                  onDelete={handleDeleteProfile}
                />
              )}
            </div>
          )}

          {/* Step 3: Run Simulation */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', marginBottom: '1rem' }}>
              <div style={{
                width: 32, height: 32, borderRadius: '50%',
                background: 'var(--gradient-primary)',
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                fontSize: '0.875rem', fontWeight: 700, color: 'white',
              }}>3</div>
              <h3>Run Simulation</h3>
            </div>
            <button
              className="btn-primary"
              onClick={handleRunSimulation}
              disabled={!canRun}
              style={{ width: '100%', padding: '1rem', fontSize: '1rem' }}
            >
              <Play size={20} />
              {isSimulating ? 'Running Simulation...' : 'Run Monte Carlo Simulation (10,000 iterations)'}
            </button>
            {!hasData && (
              <p style={{ textAlign: 'center', marginTop: '0.75rem', fontSize: '0.8125rem', color: 'var(--color-text-muted)' }}>
                Please upload both data files and select an expense profile
              </p>
            )}
          </div>

          {/* Progress */}
          {isSimulating && progress && (
            <SimulationProgress
              current={progress.current}
              total={progress.total}
              percentage={progress.percentage}
            />
          )}

          {/* Results */}
          {results.length > 0 && (
            <ResultsSummary results={results} simulationType={simulationType} />
          )}

          {/* Chart Controls */}
          {displayChart && (
            <div className="glass-card animate-fade-in" style={{ padding: '1rem 1.5rem', marginBottom: '1.5rem', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
              <div>
                <h4 style={{ margin: 0 }}>Chart Controls</h4>
                <p style={{ margin: 0, fontSize: '0.8125rem', color: 'var(--color-text-muted)' }}>Set the starting rank for the 11 highlighted scenarios (e.g., 7500 plots rank 7500 to 8500)</p>
              </div>
              
              <div style={{ display: 'flex', alignItems: 'center', gap: '1.5rem' }}>
                {simulationType === 'swr' && (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                    <label style={{ fontSize: '0.875rem', fontWeight: 600, color: 'var(--color-text-primary)' }}>Viewing SWR:</label>
                    <select 
                      className="input-field" 
                      value={activeSWR}
                      onChange={(e) => setActiveSWR(parseFloat(e.target.value))}
                      style={{ width: '100px', padding: '0.5rem', cursor: 'pointer' }}
                    >
                      {baseSWRList.map(swr => (
                        <option key={swr} value={swr}>{(swr * 100).toFixed(1)}%</option>
                      ))}
                    </select>
                  </div>
                )}
                
                <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
                  <label style={{ fontSize: '0.875rem', fontWeight: 600, color: 'var(--color-text-primary)' }}>Start Rank:</label>
                <input
                  type="number"
                  className="input-field"
                  style={{ width: '100px', padding: '0.5rem' }}
                  value={chartBaseRank}
                  onChange={(e) => setChartBaseRank(Number(e.target.value))}
                  min={1}
                  max={9990}
                  disabled={simulationType === 'swr' && yearlySummary.length === 0}
                  title={simulationType === 'swr' && yearlySummary.length === 0 ? 'Run a fresh simulation to adjust rank' : ''}
                />
                </div>
              </div>
            </div>
          )}

          {/* Chart */}
          {displayChart && (
            <CorpusSurvivalChart 
              chartData={displayChart} 
              currentAge={age} 
              withdrawalAmount={withdrawalAmount}
              withdrawalYear={withdrawalYear}
              withdrawalStartAge={withdrawalStartAge}
            />
          )}

          {/* Generate Chart from existing Excel */}
          <div className="glass-card" style={{ padding: '1.5rem' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem', marginBottom: '1rem' }}>
              <div style={{
                width: 40, height: 40, borderRadius: 'var(--radius-md)',
                background: 'rgba(245,158,11,0.1)', display: 'flex',
                alignItems: 'center', justifyContent: 'center',
              }}>
                <FileSpreadsheet size={20} color="var(--color-accent-amber)" />
              </div>
              <div>
                <h3 style={{ marginBottom: '0.25rem' }}>Generate Chart from Excel</h3>
                <p style={{ fontSize: '0.8125rem', color: 'var(--color-text-muted)' }}>
                  Upload a previously downloaded simulation Excel to generate the chart
                </p>
              </div>
            </div>
            <FileUpload
              label="Simulation Results Excel"
              description='Upload an Excel file with "Simulations Input" sheet'
              onFileLoaded={handleChartFromExcel}
              isLoaded={false}
            />
          </div>
        </div>
      </main>

      {/* View Quarterly Returns Modal */}
      {viewingQuarterly && assetReturns && (
        <div
          style={{
            position: 'fixed', inset: 0, zIndex: 1000,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            background: 'rgba(0,0,0,0.6)', backdropFilter: 'blur(4px)',
          }}
          onClick={() => setViewingQuarterly(false)}
        >
          <div
            className="glass-card-static"
            style={{
              width: '100%', maxWidth: 700, maxHeight: '80vh',
              padding: '1.5rem', overflow: 'auto',
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
              <h4>Quarterly Returns Data ({assetReturns.equity.length} records)</h4>
              <button className="btn-icon" onClick={() => setViewingQuarterly(false)}>
                <X size={16} />
              </button>
            </div>
            <table className="data-table">
              <thead>
                <tr>
                  <th>Quarter</th>
                  <th style={{ textAlign: 'right' }}>Equity</th>
                  <th style={{ textAlign: 'right' }}>Real Estate</th>
                  <th style={{ textAlign: 'right' }}>Commodity</th>
                  <th style={{ textAlign: 'right' }}>Debt</th>
                  <th style={{ textAlign: 'right' }}>Alternative</th>
                </tr>
              </thead>
              <tbody>
                {assetReturns.equity.map((eq, i) => (
                  <tr key={i}>
                    <td>Q{i + 1}</td>
                    <td style={{ textAlign: 'right' }}>{(eq * 100).toFixed(2)}%</td>
                    <td style={{ textAlign: 'right' }}>{(assetReturns.realEstate[i] * 100).toFixed(2)}%</td>
                    <td style={{ textAlign: 'right' }}>{(assetReturns.commodity[i] * 100).toFixed(2)}%</td>
                    <td style={{ textAlign: 'right' }}>{(assetReturns.debt[i] * 100).toFixed(2)}%</td>
                    <td style={{ textAlign: 'right' }}>{(assetReturns.alternative[i] * 100).toFixed(2)}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer style={{
        textAlign: 'center', padding: '1.5rem',
        borderTop: '1px solid var(--color-border)',
        fontSize: '0.75rem', color: 'var(--color-text-muted)',
      }}>
        <p>Monte Carlo Retirement Simulation • All computations run locally in your browser</p>
      </footer>
    </div>
  );
}

export default App;
