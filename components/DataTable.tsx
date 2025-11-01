
import React, { useState, useEffect, useRef } from 'react';
import type { SheetData, CellData, CellStyle } from '../types';
import { BoldIcon, ItalicIcon, AlignLeftIcon, AlignCenterIcon, AlignRightIcon, UnderlineIcon, StrikethroughIcon, TextColorIcon } from './icons';

type Selection = {
  start: { row: number, col: number };
  end: { row: number, col: number };
} | null;

interface DataTableProps {
  sheets: SheetData[];
  editable?: boolean;
  onCellChange?: (sheetIndex: number, rowIndex: number, cellIndex: number, value: string) => void;
  onStyleChange?: (sheetIndex: number, selection: Selection, style: Partial<CellStyle>) => void;
}

interface ActiveCell {
  sheetIndex: number;
  rowIndex: number;
  cellIndex: number;
}

const PRESET_COLORS = ['#111827', '#F9FAFB', '#EF4444', '#22C55E', '#3B82F6', '#F97316', '#8B5CF6'];

const FormattingToolbar: React.FC<{
  target: HTMLElement | null,
  activeCellData: CellData | null,
  onStyleChange: (style: Partial<CellStyle>) => void
}> = ({ target, activeCellData, onStyleChange }) => {
    const [showColorPicker, setShowColorPicker] = useState(false);
    const colorPickerRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (colorPickerRef.current && !colorPickerRef.current.contains(event.target as Node)) {
                setShowColorPicker(false);
            }
        };
        document.addEventListener("mousedown", handleClickOutside);
        return () => document.removeEventListener("mousedown", handleClickOutside);
    }, []);
    
    if (!target) return null;

    const rect = target.getBoundingClientRect();
    const style = activeCellData?.style || {};
    
    const buttonClass = (active: boolean) => 
      `p-1.5 rounded-md transition-colors ${active ? 'bg-primary text-white' : 'bg-secondary text-text-secondary hover:bg-border-color'}`;

    return (
        <div 
          className="absolute z-30 bg-secondary border border-border-color rounded-lg shadow-lg flex items-center p-1 gap-1"
          style={{
              top: rect.top - 50,
              left: rect.left + (rect.width / 2) - 140,
              transform: 'translateY(-10px)'
          }}
          onMouseDown={e => e.preventDefault()} // Prevent losing focus from input
        >
            <button className={buttonClass(!!style.bold)} onClick={() => onStyleChange({ bold: !style.bold })}><BoldIcon className="w-5 h-5" /></button>
            <button className={buttonClass(!!style.italic)} onClick={() => onStyleChange({ italic: !style.italic })}><ItalicIcon className="w-5 h-5" /></button>
            <button className={buttonClass(!!style.underline)} onClick={() => onStyleChange({ underline: !style.underline })}><UnderlineIcon className="w-5 h-5" /></button>
            <button className={buttonClass(!!style.strikethrough)} onClick={() => onStyleChange({ strikethrough: !style.strikethrough })}><StrikethroughIcon className="w-5 h-5" /></button>

            <div className="relative" ref={colorPickerRef}>
                <button onClick={() => setShowColorPicker(s => !s)} className="p-1.5 rounded-md text-text-secondary hover:bg-border-color">
                    <TextColorIcon className="w-5 h-5" color={style.color || '#111827'} />
                </button>
                {showColorPicker && (
                    <div className="absolute z-10 top-full mt-2 bg-secondary border border-border-color rounded-md shadow-lg p-2 grid grid-cols-4 gap-2">
                        {PRESET_COLORS.map(color => (
                            <button
                                key={color}
                                onClick={() => { onStyleChange({ color }); setShowColorPicker(false); }}
                                className="w-6 h-6 rounded-full border border-border-color"
                                style={{ backgroundColor: color }}
                                aria-label={`Set color to ${color}`}
                            />
                        ))}
                        <button
                            onClick={() => { onStyleChange({ color: undefined }); setShowColorPicker(false); }}
                            className="w-6 h-6 rounded-full border border-border-color bg-no-repeat bg-center bg-cover"
                            style={{ backgroundImage: 'linear-gradient(to top right, transparent 48%, red 48%, red 52%, transparent 52%)' }}
                            aria-label="Remove color"
                            title="Remove color"
                        />
                    </div>
                )}
            </div>

            <div className="w-px h-6 bg-border-color mx-1"></div>
            <button className={buttonClass(style.align === 'left')} onClick={() => onStyleChange({ align: 'left' })}><AlignLeftIcon className="w-5 h-5" /></button>
            <button className={buttonClass(style.align === 'center')} onClick={() => onStyleChange({ align: 'center' })}><AlignCenterIcon className="w-5 h-5" /></button>
            <button className={buttonClass(style.align === 'right')} onClick={() => onStyleChange({ align: 'right' })}><AlignRightIcon className="w-5 h-5" /></button>
        </div>
    );
};

const DataTable: React.FC<DataTableProps> = ({ sheets, editable = false, onCellChange, onStyleChange }) => {
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [activeCell, setActiveCell] = useState<ActiveCell | null>(null);
  const [selection, setSelection] = useState<Selection>(null);
  const [toolbarTarget, setToolbarTarget] = useState<HTMLElement | null>(null);
  const tableContainerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setActiveCell(null);
    setSelection(null);
  }, [activeSheetIndex]);
  
  useEffect(() => {
    if (editable && activeCell && activeCell.sheetIndex === activeSheetIndex) {
        const { rowIndex, cellIndex } = activeCell;
        const input = tableContainerRef.current?.querySelector<HTMLInputElement>(`input[data-row='${rowIndex}'][data-cell='${cellIndex}']`);
        if (input) {
            input.focus();
            input.select();
            setToolbarTarget(input.parentElement);
        } else {
            setToolbarTarget(null);
        }
    } else {
        setToolbarTarget(null);
    }
  }, [activeCell, editable, activeSheetIndex]);

  if (!sheets || sheets.length === 0) {
    return <p className="text-center text-text-secondary">No data to display.</p>;
  }

  const activeSheet = sheets[activeSheetIndex];
  if (!activeSheet || !activeSheet.data) {
     return <p className="text-center text-text-secondary">Selected sheet has no data.</p>;
  }
  
  const headers = activeSheet.data[0] || [];
  const rows = activeSheet.data.slice(1);
  const totalRows = activeSheet.data.length;
  const totalCols = headers.length;

  const handleCellClick = (e: React.MouseEvent, rowIndex: number, cellIndex: number) => {
      const newActiveCell = { sheetIndex: activeSheetIndex, rowIndex, cellIndex };
      if (e.shiftKey && activeCell) {
          setSelection({
              start: { row: Math.min(activeCell.rowIndex, rowIndex), col: Math.min(activeCell.cellIndex, cellIndex) },
              end: { row: Math.max(activeCell.rowIndex, rowIndex), col: Math.max(activeCell.cellIndex, cellIndex) },
          });
      } else {
          setActiveCell(newActiveCell);
          setSelection({ start: { row: rowIndex, col: cellIndex }, end: { row: rowIndex, col: cellIndex }});
      }
  };

  const handleKeyDown = (e: React.KeyboardEvent, rowIndex: number, cellIndex: number) => {
    if (!['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Tab'].includes(e.key)) return;
    e.preventDefault();

    let nextRow = rowIndex, nextCell = cellIndex;

    switch (e.key) {
      case 'ArrowUp': nextRow = Math.max(0, rowIndex - 1); break;
      case 'ArrowDown': nextRow = Math.min(totalRows - 1, rowIndex + 1); break;
      case 'ArrowLeft': nextCell = Math.max(0, cellIndex - 1); break;
      case 'ArrowRight': nextCell = Math.min(totalCols - 1, cellIndex + 1); break;
      case 'Tab':
        if (e.shiftKey) {
            if (cellIndex > 0) nextCell = cellIndex - 1;
            else if (rowIndex > 0) { nextRow = rowIndex - 1; nextCell = totalCols - 1; }
        } else {
            if (cellIndex < totalCols - 1) nextCell = cellIndex + 1;
            else if (rowIndex < totalRows - 1) { nextRow = rowIndex + 1; nextCell = 0; }
        }
        break;
    }

    if (nextRow !== rowIndex || nextCell !== cellIndex) {
        const newActiveCell = { sheetIndex: activeSheetIndex, rowIndex: nextRow, cellIndex: nextCell };
        setActiveCell(newActiveCell);
        setSelection({ start: { row: nextRow, col: nextCell }, end: { row: nextRow, col: nextCell }});
    }
  };

  const isCellSelected = (rowIndex: number, cellIndex: number): boolean => {
      if (!selection) return false;
      return rowIndex >= selection.start.row && rowIndex <= selection.end.row &&
             cellIndex >= selection.start.col && cellIndex <= selection.end.col;
  }

  const getCellStyle = (cell: CellData): React.CSSProperties => {
      const style: React.CSSProperties = {};
      if (cell.style?.bold) style.fontWeight = 'bold';
      if (cell.style?.italic) style.fontStyle = 'italic';
      if (cell.style?.align) style.textAlign = cell.style.align;
      if (cell.style?.color) style.color = cell.style.color;
      
      const decorations = [];
      if (cell.style?.underline) decorations.push('underline');
      if (cell.style?.strikethrough) decorations.push('line-through');
      if (decorations.length > 0) {
        style.textDecorationLine = decorations.join(' ');
      }

      return style;
  };
  
  const activeCellData = activeCell ? sheets[activeCell.sheetIndex]?.data[activeCell.rowIndex]?.[activeCell.cellIndex] : null;

  return (
    <div className="w-full bg-secondary rounded-lg border border-border-color shadow-lg overflow-hidden h-full flex flex-col">
      {editable && <FormattingToolbar target={toolbarTarget} activeCellData={activeCellData} onStyleChange={(style) => onStyleChange?.(activeSheetIndex, selection, style)} />}
      <div className="p-4 border-b border-border-color">
        <div className="flex space-x-2 border-b border-border-color pb-2 mb-2 overflow-x-auto">
          {sheets.map((sheet, index) => (
            <button key={index} onClick={() => setActiveSheetIndex(index)} className={`px-4 py-2 text-sm font-medium rounded-md whitespace-nowrap transition-colors duration-200 ${ activeSheetIndex === index ? 'bg-primary text-white' : 'bg-background text-text-secondary hover:bg-border-color hover:text-text-main' }`} >
              {sheet.sheetName}
            </button>
          ))}
        </div>
        <h3 className="text-lg font-semibold text-text-main">{activeSheet.sheetName}</h3>
        <p className="text-sm text-text-secondary">{rows.length} rows extracted</p>
      </div>
      <div ref={tableContainerRef} className="overflow-auto flex-grow relative">
        <table className="w-full text-sm text-left text-text-secondary table-fixed">
          <thead className="text-xs text-text-main uppercase bg-border-color sticky top-0 z-10">
            <tr>
              {headers.map((headerCell, index) => (
                <th key={index} scope="col" className="p-0 whitespace-nowrap border-l border-b border-border-color first:border-l-0" style={getCellStyle(headerCell)}>
                  <div className={`relative px-6 py-3 ${isCellSelected(0, index) ? 'bg-primary/20' : ''}`} onClick={(e) => editable && handleCellClick(e, 0, index)}>
                    {editable && activeCell?.rowIndex === 0 && activeCell?.cellIndex === index ? (
                      <input type="text" value={headerCell.value} onChange={(e) => onCellChange?.(activeSheetIndex, 0, index, e.target.value)} onKeyDown={(e) => handleKeyDown(e, 0, index)} className="w-full bg-transparent outline-none font-bold" data-row={0} data-cell={index}/>
                    ) : ( headerCell.value )}
                  </div>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, rowIndex) => (
              <tr key={rowIndex} className="bg-secondary last:border-b-0">
                {row.map((cell, cellIndex) => (
                  <td key={cellIndex} className={`p-0 whitespace-nowrap border-t border-l border-border-color first:border-l-0`}>
                    <div className={`relative px-6 py-4 ${isCellSelected(rowIndex + 1, cellIndex) ? 'bg-primary/20' : ''}`} style={getCellStyle(cell)} onClick={(e) => editable && handleCellClick(e, rowIndex + 1, cellIndex)}>
                        {editable && activeCell?.rowIndex === rowIndex + 1 && activeCell?.cellIndex === cellIndex ? (
                            <input type="text" value={cell.value} onChange={(e) => onCellChange?.(activeSheetIndex, rowIndex + 1, cellIndex, e.target.value)} onKeyDown={(e) => handleKeyDown(e, rowIndex + 1, cellIndex)} className="w-full bg-transparent outline-none" data-row={rowIndex + 1} data-cell={cellIndex} style={{all: 'unset', width: '100%'}}/>
                        ) : ( cell.value )}
                    </div>
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default DataTable;