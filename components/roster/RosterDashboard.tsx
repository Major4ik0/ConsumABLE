import React, { useState, useEffect, useMemo } from 'react';
import {
  Cadet,
  AbsenceRecord,
  TriSliceBalance,
  RosterSliceBalance,
  StagingCadetDTO
} from './types';
import { ImportPreviewModal } from './ImportPreviewModal';

type SliceTab = 'total' | 'male' | 'female';

interface RosterDashboardProps {
  initialDate?: string;
  apiBaseUrl?: string;
}

export const RosterDashboard: React.FC<RosterDashboardProps> = ({
  initialDate,
  apiBaseUrl = ''
}) => {
  const [selectedDate, setSelectedDate] = useState<string>(
    initialDate || new Date().toISOString().split('T')[0]
  );
  const [activeSlice, setActiveSlice] = useState<SliceTab>('total');
  const [cadets, setCadets] = useState<Cadet[]>([]);
  const [absences, setAbsences] = useState<AbsenceRecord[]>([]);
  const [categories, setCategories] = useState<string[]>([]);
  const [balance, setBalance] = useState<TriSliceBalance | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);

  // Staging Wizard modal state
  const [isWizardOpen, setIsWizardOpen] = useState<boolean>(false);
  const [stagingData, setStagingData] = useState<StagingCadetDTO[]>([]);

  // Fetch daily data from API
  const loadDailyData = async (dateStr: string) => {
    setIsLoading(true);
    try {
      const res = await fetch(`${apiBaseUrl}/api/roster/daily?report_date=${dateStr}`);
      if (res.ok) {
        const data = await res.json();
        setCadets(data.cadets || []);
        setAbsences(data.absences || []);
        setCategories(data.categories || []);
        setBalance(data.balance || null);
      }
    } catch (err) {
      console.error('Ошибка загрузки данных расхода:', err);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    loadDailyData(selectedDate);
  }, [selectedDate]);

  // Handle Excel upload -> triggers parse-preview endpoint -> opens Staging Wizard
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    try {
      setIsLoading(true);
      const res = await fetch(`${apiBaseUrl}/api/roster/parse-preview`, {
        method: 'POST',
        body: formData
      });
      if (res.ok) {
        const data = await res.json();
        setStagingData(data.cadets);
        setIsWizardOpen(true);
      } else {
        const err = await res.json();
        alert(`Ошибка парсинга: ${err.detail || 'Не удалось прочитать файл'}`);
      }
    } catch (err) {
      alert(`Ошибка сети: ${err}`);
    } finally {
      setIsLoading(false);
      e.target.value = '';
    }
  };

  // Commit corrected data from Staging Wizard
  const handleCommitWizard = async (committedCadets: StagingCadetDTO[]) => {
    try {
      setIsLoading(true);
      const res = await fetch(`${apiBaseUrl}/api/roster/commit`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ cadets: committedCadets })
      });
      if (res.ok) {
        setIsWizardOpen(false);
        await loadDailyData(selectedDate);
        alert(`Личный состав успешно обновлен (${committedCadets.length} чел.)!`);
      } else {
        const err = await res.json();
        alert(`Ошибка сохранения: ${err.detail}`);
      }
    } catch (err) {
      alert(`Ошибка сети: ${err}`);
    } finally {
      setIsLoading(false);
    }
  };

  // Filter absences by active slice tab
  const filteredAbsences = useMemo(() => {
    if (activeSlice === 'male') {
      return absences.filter(a => {
        const g = String(a.gender).toLowerCase();
        return g === 'male' || g === 'м';
      });
    }
    if (activeSlice === 'female') {
      return absences.filter(a => {
        const g = String(a.gender).toLowerCase();
        return g === 'female' || g === 'ж';
      });
    }
    return absences;
  }, [absences, activeSlice]);

  // Current active slice metrics
  const currentSliceBalance: RosterSliceBalance | undefined = useMemo(() => {
    if (!balance) return undefined;
    return balance[activeSlice];
  }, [balance, activeSlice]);

  return (
    <div style={{ maxWidth: 1200, margin: '0 auto', padding: 20, fontFamily: 'sans-serif' }}>
      {/* Top Header Card */}
      <div style={cardStyle}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 14 }}>
          <div>
            <h1 style={{ margin: 0, fontSize: 22, color: '#0F172A', fontWeight: 800 }}>
              ConsumABLE — Строевой учет
            </h1>
            <p style={{ margin: '4px 0 0 0', fontSize: 13, color: '#64748B' }}>
              Автоматизация расхода личного состава с 3 срезами и контролем баланса
            </p>
          </div>

          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <label style={{ fontSize: 13, fontWeight: 600, color: '#475569' }}>Дата:</label>
            <input
              type="date"
              style={inputStyle}
              value={selectedDate}
              onChange={e => setSelectedDate(e.target.value)}
            />

            <input
              type="file"
              id="excel-upload-input"
              accept=".xlsx,.xlsm"
              style={{ display: 'none' }}
              onChange={handleFileUpload}
            />
            <button
              style={{ ...btnStyle, backgroundColor: '#4F46E5' }}
              onClick={() => document.getElementById('excel-upload-input')?.click()}
            >
              📥 Импорт Excel (Wizard)
            </button>

            <a
              href={`${apiBaseUrl}/api/roster/export-excel?report_date=${selectedDate}`}
              download
              style={{ ...btnStyle, backgroundColor: '#059669', textDecoration: 'none' }}
            >
              📊 Скачать Excel (.xlsx)
            </a>
          </div>
        </div>
      </div>

      {/* Segmented Control Tabs for 3 Slices */}
      <div style={{ margin: '20px 0 14px 0', display: 'flex', gap: 8 }}>
        <button
          style={{
            ...tabBtnStyle,
            backgroundColor: activeSlice === 'total' ? '#2563EB' : '#F1F5F9',
            color: activeSlice === 'total' ? '#FFFFFF' : '#475569',
            borderColor: activeSlice === 'total' ? '#2563EB' : '#E2E8F0'
          }}
          onClick={() => setActiveSlice('total')}
        >
          👥 Общий расход {balance ? `(${balance.total.listed_count} чел.)` : ''}
        </button>

        <button
          style={{
            ...tabBtnStyle,
            backgroundColor: activeSlice === 'male' ? '#1D4ED8' : '#F1F5F9',
            color: activeSlice === 'male' ? '#FFFFFF' : '#475569',
            borderColor: activeSlice === 'male' ? '#1D4ED8' : '#E2E8F0'
          }}
          onClick={() => setActiveSlice('male')}
        >
          👨 Мужской состав {balance ? `(${balance.male.listed_count})` : ''}
        </button>

        <button
          style={{
            ...tabBtnStyle,
            backgroundColor: activeSlice === 'female' ? '#BE185D' : '#F1F5F9',
            color: activeSlice === 'female' ? '#FFFFFF' : '#475569',
            borderColor: activeSlice === 'female' ? '#BE185D' : '#E2E8F0'
          }}
          onClick={() => setActiveSlice('female')}
        >
          👩 Женский состав {balance ? `(${balance.female.listed_count})` : ''}
        </button>
      </div>

      {/* Dynamic KPI Cards for the selected slice */}
      <div style={kpiGridStyle}>
        <div style={kpiCardStyle}>
          <div style={kpiLabelStyle}>По списку (П/С)</div>
          <div style={kpiValueStyle}>{currentSliceBalance?.listed_count ?? 0} чел.</div>
        </div>

        <div style={{ ...kpiCardStyle, borderLeft: '4px solid #10B981' }}>
          <div style={{ ...kpiLabelStyle, color: '#047857' }}>В строю (Н/Л)</div>
          <div style={{ ...kpiValueStyle, color: '#065F46' }}>{currentSliceBalance?.present_count ?? 0} чел.</div>
        </div>

        <div style={{ ...kpiCardStyle, borderLeft: '4px solid #EF4444' }}>
          <div style={{ ...kpiLabelStyle, color: '#B91C1C' }}>В расходе (Всего)</div>
          <div style={{ ...kpiValueStyle, color: '#991B1B' }}>{currentSliceBalance?.absent_count ?? 0} чел.</div>
        </div>
      </div>

      {/* Absence Breakdown by Category */}
      {currentSliceBalance && (
        <div style={{ ...cardStyle, marginTop: 14, padding: 14 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: '#475569', marginBottom: 8 }}>
            Распределение по категориям отсутствия:
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {Object.entries(currentSliceBalance.categories_breakdown).map(([cat, count]) => (
              <span key={cat} style={catBadgeStyle}>
                <strong>{cat}:</strong> {count}
              </span>
            ))}
          </div>
        </div>
      )}

      {/* Mathematical Balance Status Banner */}
      {balance && (
        <div
          style={{
            ...balanceBannerStyle,
            backgroundColor: balance.is_sum_balanced ? '#ECFDF5' : '#FEF2F2',
            color: balance.is_sum_balanced ? '#065F46' : '#991B1B',
            borderColor: balance.is_sum_balanced ? '#A7F3D0' : '#FECACA'
          }}
        >
          {balance.is_sum_balanced ? (
            <span>
              ⚖️ <strong>Сквозной математический баланс подтвержден:</strong> Общий [П/С: {balance.total.listed_count}, В строю: {balance.total.present_count}, Расход: {balance.total.absent_count}]
              = Мужчины [{balance.male.listed_count}/{balance.male.present_count}/{balance.male.absent_count}]
              + Женщины [{balance.female.listed_count}/{balance.female.present_count}/{balance.female.absent_count}] ✓
            </span>
          ) : (
            <span>
              ❌ <strong>Сбой математического баланса:</strong> {balance.error_message}
            </span>
          )}
        </div>
      )}

      {/* Table of Absences for Active Slice */}
      <div style={{ ...cardStyle, marginTop: 16, padding: 0, overflow: 'hidden' }}>
        <div style={{ padding: '14px 18px', borderBottom: '1px solid #E2E8F0', fontWeight: 700, fontSize: 14 }}>
          Список отсутствующих: {activeSlice === 'total' ? 'Все военнослужащие' : activeSlice === 'male' ? 'Мужской состав' : 'Женский состав'} ({filteredAbsences.length} чел.)
        </div>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ backgroundColor: '#F8FAFC', borderBottom: '2px solid #E2E8F0' }}>
              <th style={thStyle}>№</th>
              <th style={thStyle}>Звание</th>
              <th style={thStyle}>ФИО</th>
              <th style={thStyle}>Группа</th>
              <th style={thStyle}>Курс</th>
              <th style={thStyle}>Причина (категория)</th>
              <th style={thStyle}>Примечание</th>
            </tr>
          </thead>
          <tbody>
            {filteredAbsences.length === 0 ? (
              <tr>
                <td colSpan={7} style={{ padding: 24, textAlign: 'center', color: '#94A3B8' }}>
                  Все курсанты данного среза находятся в строю
                </td>
              </tr>
            ) : (
              filteredAbsences.map((a, idx) => (
                <tr key={a.id || idx} style={{ borderBottom: '1px solid #F1F5F9' }}>
                  <td style={{ ...tdStyle, textAlign: 'center', width: 40 }}>{idx + 1}</td>
                  <td style={tdStyle}>{a.rank}</td>
                  <td style={{ ...tdStyle, fontWeight: 600 }}>{a.short_fio || a.fio}</td>
                  <td style={tdStyle}>{a.group_num}</td>
                  <td style={{ ...tdStyle, textAlign: 'center' }}>{a.course || 1}</td>
                  <td style={tdStyle}>
                    <span style={pillBadgeStyle}>{a.type_name} ({a.category})</span>
                  </td>
                  <td style={tdStyle}>{a.note || '—'}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Staging Wizard Modal */}
      <ImportPreviewModal
        isOpen={isWizardOpen}
        initialData={stagingData}
        onClose={() => setIsWizardOpen(false)}
        onConfirm={handleCommitWizard}
      />
    </div>
  );
};

// Styles
const cardStyle: React.CSSProperties = {
  backgroundColor: '#FFFFFF',
  borderRadius: 10,
  border: '1px solid #E2E8F0',
  padding: 18,
  boxShadow: '0 1px 3px rgba(0, 0, 0, 0.05)'
};

const inputStyle: React.CSSProperties = {
  padding: '6px 12px',
  borderRadius: 6,
  border: '1px solid #CBD5E1',
  fontSize: 13,
  outline: 'none'
};

const btnStyle: React.CSSProperties = {
  padding: '8px 16px',
  borderRadius: 6,
  border: 'none',
  color: '#FFFFFF',
  fontWeight: 600,
  fontSize: 13,
  cursor: 'pointer'
};

const tabBtnStyle: React.CSSProperties = {
  padding: '10px 18px',
  borderRadius: 8,
  border: '1px solid',
  fontWeight: 700,
  fontSize: 13,
  cursor: 'pointer',
  transition: 'all 0.15s ease'
};

const kpiGridStyle: React.CSSProperties = {
  display: 'grid',
  gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
  gap: 14,
  marginTop: 14
};

const kpiCardStyle: React.CSSProperties = {
  backgroundColor: '#FFFFFF',
  border: '1px solid #E2E8F0',
  borderRadius: 8,
  padding: 16,
  boxShadow: '0 1px 2px rgba(0,0,0,0.03)'
};

const kpiLabelStyle: React.CSSProperties = {
  fontSize: 12,
  fontWeight: 700,
  color: '#64748B',
  textTransform: 'uppercase'
};

const kpiValueStyle: React.CSSProperties = {
  fontSize: 24,
  fontWeight: 800,
  color: '#0F172A',
  marginTop: 4
};

const catBadgeStyle: React.CSSProperties = {
  padding: '4px 10px',
  borderRadius: 6,
  backgroundColor: '#F1F5F9',
  fontSize: 12,
  color: '#334155'
};

const balanceBannerStyle: React.CSSProperties = {
  marginTop: 14,
  padding: '12px 16px',
  borderRadius: 8,
  border: '1px solid',
  fontSize: 13
};

const thStyle: React.CSSProperties = {
  padding: '10px 14px',
  textAlign: 'left',
  fontWeight: 700,
  color: '#475569',
  fontSize: 12
};

const tdStyle: React.CSSProperties = {
  padding: '10px 14px',
  color: '#1E293B',
  fontSize: 13
};

const pillBadgeStyle: React.CSSProperties = {
  display: 'inline-block',
  padding: '2px 8px',
  borderRadius: 12,
  fontSize: 12,
  fontWeight: 600,
  backgroundColor: '#EFF6FF',
  color: '#1E40AF'
};
