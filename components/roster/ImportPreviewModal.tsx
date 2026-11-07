import React, { useState } from 'react';
import { StagingCadetDTO, ParsePreviewResponse } from './types';

interface ImportPreviewModalProps {
  isOpen: boolean;
  initialData: StagingCadetDTO[];
  onClose: () => void;
  onConfirm: (cadets: StagingCadetDTO[]) => void;
}

export const ImportPreviewModal: React.FC<ImportPreviewModalProps> = ({
  isOpen,
  initialData,
  onClose,
  onConfirm
}) => {
  const [cadets, setCadets] = useState<StagingCadetDTO[]>(initialData);
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [targetCourse, setTargetCourse] = useState<number>(1);

  if (!isOpen) return null;

  // 1. Быстрый сдвиг курса (+1 курс всем)
  const handleIncrementCourseAll = () => {
    setCadets(prev =>
      prev.map(c => {
        if (c.status === 'active') {
          const nextCourse = c.course + 1;
          const isGraduated = nextCourse > 5;
          const warnings = [...c.warnings];
          if (isGraduated && !warnings.some(w => w.includes('выпускник'))) {
            warnings.push(`Выпускник (курс ${nextCourse} > 5)`);
          }
          return {
            ...c,
            course: nextCourse,
            status: isGraduated ? 'graduated' : 'active',
            warnings
          };
        }
        return c;
      })
    );
  };

  // 2. Массовое переопределение курса
  const handleBulkSetCourse = () => {
    setCadets(prev =>
      prev.map(c => {
        if (selectedIds.size === 0 || selectedIds.has(c.temp_id)) {
          const isGraduated = targetCourse > 5;
          const filteredWarnings = c.warnings.filter(w => !w.includes('выпускник'));
          if (isGraduated) {
            filteredWarnings.push(`Курс ${targetCourse} превышает 5 (выпускник)`);
          }
          return {
            ...c,
            course: targetCourse,
            status: isGraduated ? 'graduated' : 'active',
            warnings: filteredWarnings
          };
        }
        return c;
      })
    );
  };

  // 3. Исключить выпускников
  const handleRemoveGraduates = () => {
    setCadets(prev => prev.filter(c => c.status !== 'graduated' && c.course <= 5));
    setSelectedIds(new Set());
  };

  // 4. Переключить пол для выделенных или конкретного курсанта
  const handleToggleGender = (tempId?: string) => {
    setCadets(prev =>
      prev.map(c => {
        const shouldToggle = tempId ? c.temp_id === tempId : selectedIds.has(c.temp_id);
        if (shouldToggle) {
          const newGender = c.gender === 'male' ? 'female' : 'male';
          return {
            ...c,
            gender: newGender,
            needs_manual_review: false,
            warnings: c.warnings.filter(w => !w.includes('пол определен'))
          };
        }
        return c;
      })
    );
  };

  // Inline course change
  const handleCourseChange = (tempId: string, newCourse: number) => {
    setCadets(prev =>
      prev.map(c => {
        if (c.temp_id === tempId) {
          const isGraduated = newCourse > 5;
          const filteredWarnings = c.warnings.filter(w => !w.includes('выпускник'));
          if (isGraduated) {
            filteredWarnings.push(`Курс ${newCourse} > 5 (выпускник)`);
          }
          return {
            ...c,
            course: newCourse,
            status: isGraduated ? 'graduated' : 'active',
            warnings: filteredWarnings
          };
        }
        return c;
      })
    );
  };

  // Selection handlers
  const handleToggleSelect = (tempId: string) => {
    const next = new Set(selectedIds);
    if (next.has(tempId)) {
      next.delete(tempId);
    } else {
      next.add(tempId);
    }
    setSelectedIds(next);
  };

  const handleToggleSelectAll = () => {
    if (selectedIds.size === cadets.length) {
      setSelectedIds(new Set());
    } else {
      setSelectedIds(new Set(cadets.map(c => c.temp_id)));
    }
  };

  // Summaries
  const totalCount = cadets.length;
  const graduatedCount = cadets.filter(c => c.status === 'graduated' || c.course > 5).length;
  const malesCount = cadets.filter(c => c.gender === 'male').length;
  const femalesCount = cadets.filter(c => c.gender === 'female').length;
  const reviewCount = cadets.filter(c => c.needs_manual_review).length;

  return (
    <div style={overlayStyle}>
      <div style={modalStyle}>
        {/* Header */}
        <div style={headerStyle}>
          <div>
            <h2 style={{ margin: 0, fontSize: 18, color: '#0F172A', fontWeight: 700 }}>
              Интерактивный предпросмотр импорта (Staging Wizard)
            </h2>
            <p style={{ margin: '4px 0 0 0', fontSize: 13, color: '#64748B' }}>
              Проверьте автоопределение пола и курсы перед внесением в строевой расчет
            </p>
          </div>
          <div style={badgeContainerStyle}>
            <span style={badgeStyle}>Всего: {totalCount} чел.</span>
            <span style={{ ...badgeStyle, backgroundColor: '#EFF6FF', color: '#1D4ED8' }}>М: {malesCount}</span>
            <span style={{ ...badgeStyle, backgroundColor: '#FDF2F8', color: '#BE185D' }}>Ж: {femalesCount}</span>
            {graduatedCount > 0 && (
              <span style={{ ...badgeStyle, backgroundColor: '#FEF3C7', color: '#B45309' }}>
                Выпускников: {graduatedCount}
              </span>
            )}
            {reviewCount > 0 && (
              <span style={{ ...badgeStyle, backgroundColor: '#FFFBEB', color: '#D97706' }}>
                Проверить пол: {reviewCount}
              </span>
            )}
          </div>
        </div>

        {/* Action Toolbar */}
        <div style={toolbarStyle}>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <button style={primaryBtnStyle} onClick={handleIncrementCourseAll}>
              +1 курс всем (Перевод на след. год)
            </button>

            <span style={{ color: '#CBD5E1' }}>|</span>

            <span style={{ fontSize: 13, fontWeight: 600, color: '#475569' }}>Назначить курс:</span>
            <select
              style={selectStyle}
              value={targetCourse}
              onChange={e => setTargetCourse(Number(e.target.value))}
            >
              {[1, 2, 3, 4, 5].map(n => (
                <option key={n} value={n}>{n} курс</option>
              ))}
            </select>
            <button style={secondaryBtnStyle} onClick={handleBulkSetCourse}>
              Применить к {selectedIds.size > 0 ? `выделенным (${selectedIds.size})` : 'всем'}
            </button>

            <span style={{ color: '#CBD5E1' }}>|</span>

            {graduatedCount > 0 && (
              <button style={warningBtnStyle} onClick={handleRemoveGraduates}>
                Исключить выпускников ({graduatedCount})
              </button>
            )}

            <button style={purpleBtnStyle} onClick={() => handleToggleGender()}>
              Переключить пол (М/Ж)
            </button>
          </div>
        </div>

        {/* Interactive Table */}
        <div style={tableContainerStyle}>
          <table style={tableStyle}>
            <thead>
              <tr style={{ backgroundColor: '#F8FAFC', borderBottom: '2px solid #E2E8F0' }}>
                <th style={{ ...thStyle, width: 40, textAlign: 'center' }}>
                  <input
                    type="checkbox"
                    checked={selectedIds.size === cadets.length && cadets.length > 0}
                    onChange={handleToggleSelectAll}
                  />
                </th>
                <th style={{ ...thStyle, width: 40, textAlign: 'center' }}>№</th>
                <th style={{ ...thStyle, width: 100 }}>Звание</th>
                <th style={thStyle}>ФИО</th>
                <th style={{ ...thStyle, width: 90, textAlign: 'center' }}>Группа</th>
                <th style={{ ...thStyle, width: 110, textAlign: 'center' }}>Курс (редакт.)</th>
                <th style={{ ...thStyle, width: 90, textAlign: 'center' }}>Пол</th>
                <th style={{ ...thStyle, width: 110, textAlign: 'center' }}>Статус</th>
                <th style={thStyle}>Предупреждения / Проверка</th>
              </tr>
            </thead>
            <tbody>
              {cadets.map((c, idx) => {
                const isSelected = selectedIds.has(c.temp_id);
                const isGraduated = c.status === 'graduated' || c.course > 5;
                const isFemale = c.gender === 'female';

                return (
                  <tr
                    key={c.temp_id}
                    style={{
                      backgroundColor: isGraduated ? '#FFFBEB' : isSelected ? '#EFF6FF' : idx % 2 === 0 ? '#FFFFFF' : '#F8FAFC',
                      borderBottom: '1px solid #F1F5F9'
                    }}
                  >
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      <input
                        type="checkbox"
                        checked={isSelected}
                        onChange={() => handleToggleSelect(c.temp_id)}
                      />
                    </td>
                    <td style={{ ...tdStyle, textAlign: 'center', color: '#64748B' }}>{idx + 1}</td>
                    <td style={tdStyle}>{c.rank}</td>
                    <td style={{ ...tdStyle, fontWeight: 600 }}>{c.fio}</td>
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      <span style={pillBadgeStyle}>{c.group_num}</span>
                    </td>
                    {/* Inline Course Selector */}
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      <select
                        style={{ ...selectStyle, padding: '3px 8px', fontSize: 12 }}
                        value={c.course}
                        onChange={e => handleCourseChange(c.temp_id, Number(e.target.value))}
                      >
                        {[1, 2, 3, 4, 5, 6].map(cr => (
                          <option key={cr} value={cr}>{cr} курс</option>
                        ))}
                      </select>
                    </td>
                    {/* Inline Gender Toggle */}
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      <button
                        title="Нажмите для смены пола"
                        style={{
                          ...genderToggleBtnStyle,
                          backgroundColor: isFemale ? '#FDF2F8' : '#EFF6FF',
                          color: isFemale ? '#BE185D' : '#1D4ED8',
                          borderColor: isFemale ? '#FBCFE8' : '#BFDBFE'
                        }}
                        onClick={() => handleToggleGender(c.temp_id)}
                      >
                        {isFemale ? 'Ж' : 'М'}
                      </button>
                    </td>
                    {/* Status */}
                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                      {isGraduated ? (
                        <span style={{ ...pillBadgeStyle, backgroundColor: '#FEF3C7', color: '#92400E' }}>
                          Выпускник
                        </span>
                      ) : (
                        <span style={{ ...pillBadgeStyle, backgroundColor: '#ECFDF5', color: '#065F46' }}>
                          В строю
                        </span>
                      )}
                    </td>
                    {/* Warnings */}
                    <td style={tdStyle}>
                      {c.warnings.length > 0 ? (
                        <span style={{ color: '#DC2626', fontSize: 12, fontWeight: 500 }}>
                          {c.warnings.join('; ')}
                        </span>
                      ) : (
                        <span style={{ color: '#16A34A', fontSize: 12 }}>Норма</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Footer Actions */}
        <div style={footerStyle}>
          <div style={{ fontSize: 12, color: '#64748B' }}>
            <i>Вы можете менять курс и пол прямо в ячейках таблицы перед сохранением.</i>
          </div>
          <div style={{ display: 'flex', gap: 10 }}>
            <button style={secondaryBtnStyle} onClick={onClose}>
              Отмена
            </button>
            <button
              style={{ ...primaryBtnStyle, backgroundColor: '#059669' }}
              onClick={() => onConfirm(cadets)}
            >
              Подтвердить и внести в расход ({cadets.length} чел.)
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

// Inline Styles
const overlayStyle: React.CSSProperties = {
  position: 'fixed',
  top: 0,
  left: 0,
  right: 0,
  bottom: 0,
  backgroundColor: 'rgba(15, 23, 42, 0.65)',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'center',
  zIndex: 1000,
  padding: 20
};

const modalStyle: React.CSSProperties = {
  backgroundColor: '#FFFFFF',
  borderRadius: 12,
  width: '100%',
  maxWidth: 1180,
  height: '90vh',
  maxHeight: 820,
  display: 'flex',
  flexDirection: 'column',
  boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.1)',
  overflow: 'hidden'
};

const headerStyle: React.CSSProperties = {
  padding: '16px 20px',
  borderBottom: '1px solid #E2E8F0',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  flexWrap: 'wrap',
  gap: 12
};

const badgeContainerStyle: React.CSSProperties = {
  display: 'flex',
  gap: 8,
  flexWrap: 'wrap'
};

const badgeStyle: React.CSSProperties = {
  padding: '4px 10px',
  borderRadius: 6,
  fontSize: 12,
  fontWeight: 600,
  backgroundColor: '#F1F5F9',
  color: '#334155'
};

const toolbarStyle: React.CSSProperties = {
  padding: '12px 20px',
  backgroundColor: '#F8FAFC',
  borderBottom: '1px solid #E2E8F0'
};

const tableContainerStyle: React.CSSProperties = {
  flex: 1,
  overflowY: 'auto'
};

const tableStyle: React.CSSProperties = {
  width: '100%',
  borderCollapse: 'collapse',
  fontSize: 13,
  textAlign: 'left'
};

const thStyle: React.CSSProperties = {
  padding: '10px 12px',
  fontSize: 12,
  fontWeight: 700,
  color: '#475569',
  borderBottom: '2px solid #E2E8F0'
};

const tdStyle: React.CSSProperties = {
  padding: '8px 12px',
  fontSize: 13,
  color: '#1E293B'
};

const footerStyle: React.CSSProperties = {
  padding: '14px 20px',
  borderTop: '1px solid #E2E8F0',
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between'
};

const primaryBtnStyle: React.CSSProperties = {
  backgroundColor: '#2563EB',
  color: '#FFFFFF',
  border: 'none',
  padding: '8px 16px',
  borderRadius: 6,
  fontWeight: 600,
  fontSize: 13,
  cursor: 'pointer'
};

const secondaryBtnStyle: React.CSSProperties = {
  backgroundColor: '#FFFFFF',
  color: '#334155',
  border: '1px solid #CBD5E1',
  padding: '8px 16px',
  borderRadius: 6,
  fontWeight: 600,
  fontSize: 13,
  cursor: 'pointer'
};

const warningBtnStyle: React.CSSProperties = {
  backgroundColor: '#FEF3C7',
  color: '#92400E',
  border: '1px solid #FDE68A',
  padding: '8px 14px',
  borderRadius: 6,
  fontWeight: 600,
  fontSize: 13,
  cursor: 'pointer'
};

const purpleBtnStyle: React.CSSProperties = {
  backgroundColor: '#F3E8FF',
  color: '#6B21A8',
  border: '1px solid #E9D5FF',
  padding: '8px 14px',
  borderRadius: 6,
  fontWeight: 600,
  fontSize: 13,
  cursor: 'pointer'
};

const selectStyle: React.CSSProperties = {
  padding: '6px 10px',
  borderRadius: 6,
  border: '1px solid #CBD5E1',
  fontSize: 13,
  backgroundColor: '#FFFFFF'
};

const pillBadgeStyle: React.CSSProperties = {
  display: 'inline-block',
  padding: '2px 8px',
  borderRadius: 12,
  fontSize: 12,
  fontWeight: 600,
  backgroundColor: '#F1F5F9',
  color: '#475569'
};

const genderToggleBtnStyle: React.CSSProperties = {
  padding: '3px 10px',
  borderRadius: 6,
  fontSize: 12,
  fontWeight: 700,
  border: '1px solid',
  cursor: 'pointer'
};
