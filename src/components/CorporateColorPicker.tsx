import React, { useState, useEffect } from 'react';
import { COLOR_PRESETS, getCorporateColor, setCorporateColor, hexToRgb, DEFAULT_CORPORATE_COLOR } from '../lib/theme';
import { Palette, Check, RotateCcw } from 'lucide-react';
import { Card } from './Card';
import { Button } from './Button';

export const CorporateColorPicker: React.FC<{ onColorChange?: (color: string) => void }> = ({ onColorChange }) => {
  const [activeColor, setActiveColor] = useState<string>(getCorporateColor());
  const [customInput, setCustomInput] = useState<string>(getCorporateColor());

  useEffect(() => {
    const current = getCorporateColor();
    setActiveColor(current);
    setCustomInput(current);
  }, []);

  const handleSelectColor = (hex: string) => {
    setActiveColor(hex);
    setCustomInput(hex);
    setCorporateColor(hex);
    if (onColorChange) onColorChange(hex);
  };

  const handleCustomSubmit = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    let val = customInput.trim();
    if (!val.startsWith('#')) {
      val = '#' + val;
    }
    if (/^#[0-9A-Fa-f]{6}$/.test(val) || /^#[0-9A-Fa-f]{3}$/.test(val)) {
      handleSelectColor(val);
    } else {
      alert('Por favor, informe uma cor hexadecimal válida (ex: #059669 ou #2563eb)');
    }
  };

  const handleReset = () => {
    handleSelectColor(DEFAULT_CORPORATE_COLOR);
  };

  const [r, g, b] = hexToRgb(activeColor);

  return (
    <Card className="p-6">
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 mb-4">
        <div>
          <h3 className="text-base font-black text-slate-800 uppercase tracking-wider flex items-center gap-2">
            <Palette style={{ color: activeColor }} size={20} />
            Cor Corporativa Principal & Relatórios PDF
          </h3>
          <p className="text-xs text-slate-500 font-semibold mt-1">
            Personalize a cor de destaque do sistema e os cabeçalhos/tabelas de todos os relatórios PDF gerados.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <Button
            onClick={handleReset}
            variant="outline"
            size="sm"
            className="text-xs border-slate-200 text-slate-600 hover:bg-slate-50 flex items-center gap-1 font-bold"
            title="Restaurar padrão SEBRAE"
          >
            <RotateCcw size={13} />
            Padrão SEBRAE
          </Button>
        </div>
      </div>

      {/* Preset Grid */}
      <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-3 mb-5">
        {COLOR_PRESETS.map((preset) => {
          const isSelected = activeColor.toLowerCase() === preset.hex.toLowerCase();
          return (
            <button
              key={preset.id}
              type="button"
              onClick={() => handleSelectColor(preset.hex)}
              className={`p-3 rounded-xl border text-left transition-all flex flex-col justify-between relative group cursor-pointer ${
                isSelected
                  ? 'ring-2 ring-offset-1 border-transparent shadow-sm'
                  : 'border-slate-200/80 hover:border-slate-300 bg-white'
              }`}
              style={{
                borderColor: isSelected ? preset.hex : undefined,
                backgroundColor: isSelected ? `rgba(${preset.rgb.join(',')}, 0.06)` : '#ffffff'
              }}
            >
              <div className="flex items-center justify-between w-full mb-2">
                <div
                  className="w-7 h-7 rounded-lg shadow-inner flex items-center justify-center text-white"
                  style={{ backgroundColor: preset.hex }}
                >
                  {isSelected && <Check size={14} strokeWidth={3} />}
                </div>
                <span className="text-[10px] font-mono font-bold text-slate-400">
                  {preset.hex.toUpperCase()}
                </span>
              </div>
              <div>
                <p className="text-xs font-bold text-slate-800 leading-tight">{preset.name.split(' (')[0]}</p>
                <p className="text-[10px] text-slate-400 line-clamp-1 mt-0.5">{preset.description}</p>
              </div>
            </button>
          );
        })}
      </div>

      {/* Custom Picker & Live PDF Preview */}
      <div className="p-4 bg-slate-50/80 rounded-2xl border border-slate-200 flex flex-col md:flex-row items-center justify-between gap-4">
        {/* Custom Input */}
        <div className="flex items-center gap-3 w-full md:w-auto">
          <div className="relative">
            <input
              type="color"
              value={activeColor}
              onChange={(e) => handleSelectColor(e.target.value)}
              className="w-10 h-10 rounded-xl cursor-pointer border-2 border-white shadow-sm p-0 bg-transparent overflow-hidden"
              title="Clique para escolher qualquer cor personalizada"
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-700">Cor Personalizada (HEX)</label>
            <div className="flex items-center gap-2 mt-1">
              <input
                type="text"
                value={customInput}
                onChange={(e) => setCustomInput(e.target.value)}
                placeholder="#059669"
                className="w-28 text-xs font-mono font-bold px-2.5 py-1.5 rounded-lg border border-slate-300 bg-white uppercase focus:outline-none focus:ring-2 focus:ring-slate-400"
              />
              <Button
                onClick={handleCustomSubmit}
                size="sm"
                className="text-xs py-1.5 px-3 font-bold text-white shadow-sm"
                style={{ backgroundColor: activeColor }}
              >
                Aplicar
              </Button>
            </div>
          </div>
        </div>

        {/* Live Mock PDF Preview Bar */}
        <div className="w-full md:w-auto flex-1 max-w-sm bg-white p-3 rounded-xl border border-slate-200/80 shadow-xs">
          <div className="flex items-center justify-between mb-1.5">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider">
              Pré-visualização nos PDFs
            </span>
            <span
              className="text-[10px] font-bold px-2 py-0.5 rounded-full text-white"
              style={{ backgroundColor: activeColor }}
            >
              RGB({r}, {g}, {b})
            </span>
          </div>
          <div className="h-4 rounded-md overflow-hidden flex shadow-inner" style={{ backgroundColor: activeColor }}>
            <div className="w-1/3 bg-black/10 border-r border-white/20"></div>
            <div className="w-1/3 bg-white/10"></div>
          </div>
          <p className="text-[10px] text-slate-500 mt-1.5 flex items-center justify-between">
            <span>Cabeçalhos &amp; Tabelas: <strong style={{ color: activeColor }}>Ativo</strong></span>
            <span>Bordas &amp; Tags: <strong style={{ color: activeColor }}>Sincronizado</strong></span>
          </p>
        </div>
      </div>
    </Card>
  );
};
