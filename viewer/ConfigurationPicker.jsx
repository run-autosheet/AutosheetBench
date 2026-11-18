import React from 'react';
import {Check, ChevronDown} from 'lucide-react';

// Choose which configurations the matrix shows.
export default function ConfigurationPicker({configurations, hiddenIds, onToggle, onSelectAll, onUnselectAll}) {
  const selected = configurations.filter(configuration => !hiddenIds.has(configuration.id)).length;
  return (
    <details className="group/config relative">
      <summary className="compare-orientation-btn flex cursor-pointer list-none items-center gap-1">
        Configurations {selected}/{configurations.length}
        <ChevronDown size={12}/>
      </summary>
      <div className="absolute left-0 top-full z-40 mt-1 w-[360px] rounded border border-gray-700 bg-gray-950 p-2 shadow-xl">
        <div className="mb-2 flex items-center gap-3 border-b border-gray-800 pb-2 text-[10px]">
          <button type="button" className="text-blue-400 underline-offset-2 hover:text-blue-300 hover:underline" onClick={onSelectAll}>Select all</button>
          <button type="button" className="text-gray-400 underline-offset-2 hover:text-gray-200 hover:underline" onClick={onUnselectAll}>Unselect all</button>
        </div>
        <div className="max-h-80 space-y-0.5 overflow-auto">
          {configurations.map(configuration => {
            const checked = !hiddenIds.has(configuration.id);
            return (
              <label key={configuration.id} className="flex cursor-pointer items-start gap-2 rounded px-2 py-1.5 text-[11px] text-gray-300 hover:bg-gray-800">
                <input type="checkbox" checked={checked} onChange={() => onToggle(configuration.id)} className="sr-only"/>
                <span className={`mt-0.5 flex h-3.5 w-3.5 shrink-0 items-center justify-center rounded border ${checked ? 'border-blue-500 bg-blue-600 text-white' : 'border-gray-600'}`}>
                  {checked && <Check size={10}/>}
                </span>
                <span className="min-w-0">{configuration.label}</span>
              </label>
            );
          })}
        </div>
      </div>
    </details>
  );
}
