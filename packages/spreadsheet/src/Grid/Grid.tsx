import React, {
  useCallback,
  useRef,
  useMemo,
  useEffect,
  memo,
  useImperativeHandle,
  forwardRef,
  useState,
} from "react";
import Grid, {
  RendererProps,
  useSelection,
  GridRef,
  useEditable,
  useCopyPaste,
  CellInterface,
  SelectionArea,
  ScrollCoords,
  AreaProps,
  useSizer as useAutoSizer,
  CellOverlayRenderer as cellOverlayRenderer,
  useTouch,
  useFilter,
  FilterView,
  FilterDefinition,
  OptionalCellInterface,
  CellPosition,
  SelectionPolicy,
  useTooltip,
  StylingProps,
  DefaultTooltipProps,
} from "@rowsncolumns/grid";
import { debounce, cellIdentifier } from "@rowsncolumns/grid";
import { ThemeProvider, ColorModeProvider } from "@chakra-ui/core";
import {
  COLUMN_HEADER_WIDTH,
  DEFAULT_COLUMN_WIDTH,
  ROW_HEADER_HEIGHT,
  DEFAULT_ROW_HEIGHT,
  DARK_MODE_COLOR_LIGHT,
  HEADER_BORDER_COLOR,
  CELL_BORDER_COLOR,
  number2Alpha,
  getEditorType,
  DEFAULT_CHECKBOX_VALUES,
  castToString,
} from "./../constants";
import HeaderCell from "./../HeaderCell";
import Cell from "./../Cell";
import { GridWrapper, ThemeType } from "./../styled";
import { Cells, CellConfig, SizeType, SheetID } from "../Spreadsheet";
import { Direction } from "@rowsncolumns/grid";
import { AXIS, Formatter, SELECTION_MODE } from "../types";
import Editor from "./../Editor";
import { EditorProps } from "@rowsncolumns/grid";
import { CustomEditorProps } from "../Editor/Editor";
import FilterComponent from "./../FilterComponent";
import { FILTER_ICON_DIM } from "../FilterIcon/FilterIcon";
import { ContextMenuComponentProps } from "../ContextMenu/ContextMenu";
import { LIST_ICON_DIM } from "../ListArrow/ListArrow";
import TooltipComponent, { TooltipProps } from "./../Tooltip";

const EMPTY_ARRAY: any = [];
const EMPTY_OBJECT: any = {};

export interface GridProps {
  theme?: ThemeType;
  minColumnWidth?: number;
  minRowHeight?: number;
  rowCount?: number;
  columnCount?: number;
  CellRenderer?: React.ReactType;
  HeaderCellRenderer?: React.ReactType;
  width?: number;
  height?: number;
  cells: Cells;
  onChange: (value: React.ReactText, cell: CellInterface) => void;
  onFill?: (
    activeCell: CellInterface,
    currentSelection: SelectionArea | null,
    selections: SelectionArea[]
  ) => void;
  activeCell: CellInterface | null;
  selections: SelectionArea[];
  onSheetChange: (props: any) => void;
  selectedSheet: SheetID;
  onScroll: (state: ScrollCoords) => void;
  scrollState?: ScrollCoords;
  onActiveCellChange?: (
    cell: CellInterface | null,
    value?: React.ReactText
  ) => void;
  onSelectionChange?: (
    cell: CellInterface | null,
    selections: SelectionArea[]
  ) => void;
  onActiveCellValueChange: (
    value: React.ReactText | undefined,
    activeCell: CellInterface | null
  ) => void;
  onDelete?: (activeCell: CellInterface, selections: SelectionArea[]) => void;
  formatter?: Formatter;
  onResize?: (axis: AXIS, index: number, dimension: number) => void;
  columnSizes?: SizeType;
  rowSizes?: SizeType;
  mergedCells?: AreaProps[];
  frozenRows?: number;
  frozenColumns?: number;
  onKeyDown?: (
    e: React.KeyboardEvent<HTMLDivElement>,
    cell: CellInterface | null,
    selections: SelectionArea[]
  ) => void;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  onPaste?: (
    rows: (string | null)[][],
    activeCell: CellInterface | null
  ) => void;
  onCut?: (selection: SelectionArea) => void;
  onInsertRow?: (
    cell: CellInterface | null,
    selections: SelectionArea[]
  ) => void;
  onDeleteRow?: (
    cell: CellInterface | null,
    selections: SelectionArea[]
  ) => void;
  onInsertColumn?: (
    cell: CellInterface | null,
    selections: SelectionArea[]
  ) => void;
  onDeleteColumn?: (
    cell: CellInterface | null,
    selections: SelectionArea[]
  ) => void;
  showGridLines?: boolean;
  CellEditor?: React.ReactType<CustomEditorProps>;
  selectionPolicy?: SelectionPolicy;
  selectionMode?: SELECTION_MODE;
  isLightMode?: boolean;
  filterViews?: FilterView[];
  onChangeFilter: (
    filterIndex: number,
    columnIndex: number,
    filter: FilterDefinition
  ) => void;
  scale?: number;
  selectionTopBound?: number;
  selectionLeftBound?: number;
  ContextMenu?: React.ReactType<ContextMenuComponentProps>;
  Tooltip?: React.ReactType<TooltipProps>;
  snap?: boolean;
}

export interface RowColSelection {
  rows: number[];
  cols: number[];
}

export type RefAttributeGrid = {
  ref?: React.Ref<WorkbookGridRef>;
};

export type WorkbookGridRef = {
  getNextFocusableCell: (
    cell: CellInterface,
    direction: Direction
  ) => CellInterface | null;
  setActiveCell: (cell: CellInterface | null) => void;
  setSelections: (selection: SelectionArea[]) => void;
  resetAfterIndices?: (
    coords: OptionalCellInterface,
    shouldForceUpdate?: boolean
  ) => void;
  focus: () => void;
  makeEditable: (cell: CellInterface, value?: string, focus?: boolean) => void;
  setEditorValue: (value: string, activeCell: CellInterface) => void;
  hideEditor: () => void;
  submitEditor: (
    value: string,
    activeCell: CellInterface,
    nextActiveCell?: CellInterface | null
  ) => void;
  cancelEditor: () => void;
  resizeColumns?: (indices: number[]) => void;
  resizeRows?: (indices: number[]) => void;
  getCellBounds?: (coords: CellInterface) => AreaProps;
  getScrollPosition?: () => ScrollCoords;
  getCellCoordsFromOffset?: (x: number, y: number) => CellInterface | null;
  getCellOffsetFromCoords?: (coords: CellInterface) => CellPosition;
};

export interface ContextMenuProps {
  left: number;
  top: number;
}

interface InternalRef {
  rowCount: number;
  columnCount: number;
}

/**
 * Grid component
 * @param props
 */
const SheetGrid: React.FC<GridProps & RefAttributeGrid> = memo(
  forwardRef<WorkbookGridRef, GridProps>((props, forwardedRef) => {
    const {
      theme,
      minColumnWidth = DEFAULT_COLUMN_WIDTH,
      minRowHeight = DEFAULT_ROW_HEIGHT,
      rowCount: initialRowCount = 1000,
      columnCount: initialColumnCount = 1000,
      width,
      height,
      CellRenderer = Cell,
      HeaderCellRenderer = HeaderCell,
      cells,
      onChange,
      onFill,
      onSheetChange,
      activeCell: initialActiveCell,
      selections: initialSelections = EMPTY_ARRAY as SelectionArea[],
      selectedSheet,
      onScroll,
      scrollState,
      onActiveCellChange,
      onActiveCellValueChange,
      onDelete,
      formatter,
      onResize,
      columnSizes = EMPTY_OBJECT as SizeType,
      rowSizes = EMPTY_OBJECT as SizeType,
      mergedCells,
      frozenRows: userFrozenRows = 0,
      frozenColumns: userFrozenColumns = 0,
      onKeyDown,
      hiddenColumns: userHiddenColumns = EMPTY_ARRAY as number[],
      hiddenRows: userHiddenRows = EMPTY_ARRAY as number[],
      filterViews = EMPTY_ARRAY as FilterView[],
      onPaste,
      onCut,
      onInsertRow,
      onInsertColumn,
      onDeleteColumn,
      onDeleteRow,
      showGridLines,
      CellEditor = Editor,
      selectionPolicy,
      onSelectionChange,
      selectionMode,
      isLightMode,
      onChangeFilter,
      scale,
      selectionTopBound = 1,
      selectionLeftBound = 1,
      ContextMenu,
      Tooltip = TooltipComponent,
      snap,
    } = props;

    const gridRef = useRef<GridRef | null>(null);
    const onSheetChangeRef = useRef(debounce(onSheetChange, 100));
    const rowCount = initialRowCount + 1;
    const columnCount = initialColumnCount + 1;
    /**
     * Keep track of variables in `refs` cos we bound events to `document`
     * : mousemove, mouseup
     */
    const internalRefs = useRef<InternalRef>({ rowCount, columnCount });
    const frozenRows = Math.max(1, userFrozenRows + 1);
    const frozenColumns = Math.max(1, userFrozenColumns + 1);
    const debounceScroll = useRef(debounce(onScroll, 500));
    const [
      contextMenuProps,
      setContextMenuProps,
    ] = useState<ContextMenuProps | null>(null);
    const borderStyles = useMemo((): StylingProps => {
      return filterViews.map((filter) => {
        return {
          bounds: filter.bounds,
          style: {
            strokeWidth: 1,
            stroke: "green",
          },
        };
      });
    }, [filterViews]);

    /* Update internal refs */
    useEffect(() => {
      internalRefs.current = {
        rowCount,
        columnCount,
      };
    }, [rowCount, columnCount]);

    const isCellRowHeader = useCallback((rowIndex: number) => {
      return rowIndex === 0;
    }, []);
    const isCellColumnHeader = useCallback((columnIndex: number) => {
      return columnIndex === 0;
    }, []);

    /* Cell where filter icon will appear */
    const filterHeaderCells = useMemo(() => {
      const initialValue: Record<string, number> = {};
      return filterViews.reduce((acc, filter, index) => {
        const { bounds } = filter;
        for (let i = bounds.left; i <= bounds.right; i++) {
          acc[cellIdentifier(bounds.top, i)] = index;
        }
        return acc;
      }, initialValue);
    }, [filterViews]);

    /* Filter columns */
    const columnsWithFilter = useMemo(() => {
      const initialValue: number[] = [];
      return filterViews.reduce((acc, filter) => {
        const { bounds } = filter;
        for (let i = bounds.left; i <= bounds.right; i++) {
          acc.push(i);
        }
        return acc;
      }, initialValue);
    }, [filterViews]);

    /* Filter row range */
    const rowFilterRange = useMemo(() => {
      const initialValue: Record<string, number[]> = {};
      return filterViews.reduce((acc, { bounds }, index) => {
        acc[index] = [bounds.top, bounds.bottom];
        return acc;
      }, initialValue);
    }, [filterViews]);

    /* Has filter */
    const columnHasFilter = useCallback(
      (columnIndex) => {
        return columnsWithFilter.includes(columnIndex);
      },
      [columnsWithFilter]
    );

    const rowHasFilter = useCallback(
      (rowIndex) => {
        for (const i in rowFilterRange) {
          const [min, max] = rowFilterRange[i];
          if (rowIndex >= min && rowIndex <= max) return true;
        }
        return false;
      },
      [rowFilterRange]
    );

    useImperativeHandle(forwardedRef, () => {
      return {
        getNextFocusableCell,
        setActiveCell,
        setSelections,
        focus: () => gridRef.current?.focus(),
        resetAfterIndices: gridRef.current?.resetAfterIndices,
        makeEditable,
        setEditorValue: setValue,
        hideEditor,
        submitEditor,
        cancelEditor,
        resizeColumns: gridRef.current?.resizeColumns,
        resizeRows: gridRef.current?.resizeRows,
        getCellBounds: gridRef.current?.getCellBounds,
        getScrollPosition: gridRef.current?.getScrollPosition,
        getCellCoordsFromOffset: gridRef.current?.getCellCoordsFromOffset,
        getCellOffsetFromCoords: gridRef.current?.getCellOffsetFromCoords,
      };
    });

    /**
     * Get cell value or text
     */
    const getValue = useCallback(
      (cell: CellInterface | null): CellConfig | undefined => {
        if (!cell) return void 0;
        const { rowIndex, columnIndex } = cell;
        /* Check if its header cell */
        const isRowHeader = isCellRowHeader(rowIndex);
        const isColumnHeader = isCellColumnHeader(columnIndex);
        const cellId = cellIdentifier(rowIndex, columnIndex);
        const isFilterHeader = filterHeaderCells[cellId] !== void 0;

        if (isRowHeader) {
          return { text: number2Alpha(columnIndex - 1), fontSize: 10 };
        }

        if (isColumnHeader) {
          return { text: rowIndex, fontSize: 10 };
        }

        const cellConfig = isFilterHeader
          ? { ...cells[rowIndex]?.[columnIndex], bold: true }
          : cells[rowIndex]?.[columnIndex];

        return rowIndex in cells ? cellConfig : void 0;
      },
      [cells, filterHeaderCells]
    );

    /**
     * To bypass closures
     */
    const getValueRef = useRef(getValue);
    useEffect(() => {
      getValueRef.current = getValue;
    }, [getValue]);

    const getValueText = useCallback(
      (cell) => {
        return getValue(cell)?.text;
      },
      [cells]
    );

    /**
     * Apply filter on the cells
     */
    const hiddenFilterRows = useMemo(() => {
      const rows: Record<string, true> = {};
      for (let i = 0; i < filterViews.length; i++) {
        const filterView = filterViews[i];
        const { bounds, filters } = filterView;
        for (const columnIndex in filters) {
          const { values } = filters[columnIndex];
          for (let k = bounds.top + 1; k <= bounds.bottom; k++) {
            const cell = { rowIndex: k, columnIndex: parseInt(columnIndex) };
            const cellConfig = getValue(cell);
            const value = cellConfig?.text ?? "";
            if (!values.includes(value)) {
              rows[k] = true;
            }
          }
        }
      }
      return rows;
    }, [filterViews]);

    const hiddenRows = useMemo(() => {
      return userHiddenRows.reduce((acc, item) => {
        acc[item] = true;
        return acc;
      }, hiddenFilterRows);
    }, [hiddenFilterRows, userHiddenRows]);

    const hiddenColumns = useMemo(() => {
      const initialValue: Record<string, boolean> = {};
      return userHiddenColumns.reduce((acc, item) => {
        acc[item] = true;
        return acc;
      }, initialValue);
    }, [userHiddenColumns]);

    const isHiddenRow = useCallback(
      (rowIndex: number) => {
        return hiddenRows[rowIndex];
      },
      [hiddenRows, selectionMode]
    );
    const isHiddenColumn = useCallback(
      (columnIndex: number) => {
        return hiddenColumns[columnIndex];
      },
      [hiddenColumns]
    );

    /* Enable touch */
    const {
      isTouchDevice,
      scrollTo: scrollToTouch,
      scrollToTop: scrollToTopTouch,
    } = useTouch({
      gridRef,
    });

    /**
     * Column resizer
     */
    const { getColumnWidth, onViewChange } = useAutoSizer({
      gridRef,
      frozenRows,
      scale,
      minColumnWidth: 10,
      isHiddenRow,
      isHiddenColumn,
      getValue: (cell: CellInterface) => {
        const cellConfig = getValue(cell);
        const formattedValue = formatter
          ? formatter(cellConfig?.text, cellConfig?.datatype, cellConfig)
          : cellConfig?.text;
        const iconPadding = 5;
        const spacing =
          cellConfig?.dataValidation?.type === "list"
            ? LIST_ICON_DIM + iconPadding
            : 0;
        return { ...cellConfig, text: formattedValue, spacing };
      },
      getText: (cellConfig: CellConfig) => {
        const isFormula = cellConfig?.datatype === "formula";
        const text = isFormula ? cellConfig?.result : cellConfig?.text;
        return text !== void 0 ? castToString(text) : text;
      },
      columnSizes,
      autoResize: false,
      resizeOnScroll: false,
    });

    /* Mouse down seelction */
    const handleMouseDownSelection = useCallback(
      (
        e: React.MouseEvent<HTMLDivElement>,
        coords: CellInterface,
        startRef,
        endRef
      ) => {
        if (!gridRef.current) return;
        const { rowCount, columnCount } = internalRefs.current;
        const isShiftKey = e.nativeEvent.shiftKey;
        const isMetaKey = e.nativeEvent.ctrlKey || e.nativeEvent.metaKey;
        const isRowHeader = isCellRowHeader(coords.rowIndex);
        const isColumnHeader = isCellColumnHeader(coords.columnIndex);
        if (isRowHeader && isColumnHeader) {
          selectAll();
          return false;
        }
        if (isRowHeader || isColumnHeader) {
          const actualFrozenRows = frozenRows - 1;
          const actualFrozenColumns = frozenColumns - 1;
          const start = isColumnHeader
            ? { ...coords, columnIndex: selectionLeftBound }
            : { ...coords, rowIndex: selectionTopBound };
          const end = isColumnHeader
            ? { ...coords, columnIndex: columnCount - 1 }
            : { ...coords, rowIndex: rowCount - 1 };
          const {
            visibleRowStartIndex,
            visibleColumnStartIndex,
          } = gridRef.current.getViewPort();
          const activeCell = isColumnHeader
            ? {
                ...coords,
                columnIndex:
                  actualFrozenColumns >= selectionLeftBound
                    ? selectionLeftBound
                    : visibleColumnStartIndex,
              }
            : {
                ...coords,
                rowIndex:
                  actualFrozenRows >= selectionTopBound
                    ? selectionTopBound
                    : visibleRowStartIndex,
              };

          if (isShiftKey) {
            modifySelection(end);
          } else if (isMetaKey) {
            appendSelection(start, end);
            setActiveCellState(activeCell);
            /* Scroll to the new active cell */
            gridRef.current.scrollToItem(activeCell);
          } else {
            /* Clear all selections */
            clearSelections();
            startRef.current = start;
            endRef.current = start;
            modifySelection(end);
            setActiveCellState(activeCell);
            /* Scroll to the new active cell */
            gridRef.current.scrollToItem(activeCell);
          }
          return false;
        }
      },
      [frozenRows, frozenColumns, mergedCells]
    );

    /* Mouse move */
    const handleMouseMoveSelection = useCallback(
      (_, coords: CellInterface, startRef, endRef) => {
        const { rowCount, columnCount } = internalRefs.current;
        const isRowHeader =
          startRef.current?.rowIndex === selectionTopBound &&
          endRef.current?.rowIndex === rowCount - 1;
        const isColumnHeader =
          startRef.current?.columnIndex === selectionLeftBound &&
          endRef.current?.columnIndex === columnCount - 1;

        if (isRowHeader && isColumnHeader) return false;
        if (isRowHeader || isColumnHeader) {
          const end = isRowHeader
            ? { ...coords, rowIndex: rowCount - 1 }
            : { ...coords, columnIndex: columnCount - 1 };
          modifySelection(end);
          gridRef.current?.scrollToItem({
            rowIndex: isRowHeader ? void 0 : coords.rowIndex,
            columnIndex: isColumnHeader ? void 0 : coords.columnIndex,
          });
          return false;
        }
      },
      []
    );

    /**
     * Selection
     */
    const {
      selections,
      activeCell,
      setActiveCell,
      setActiveCellState,
      setSelections,
      modifySelection,
      newSelection,
      clearLastSelection,
      appendSelection,
      selectAll,
      clearSelections,
      ...selectionProps
    } = useSelection({
      selectionPolicy,
      selectionTopBound,
      selectionLeftBound,
      initialActiveCell,
      initialSelections,
      gridRef,
      rowCount,
      columnCount,
      onFill,
      isHiddenRow,
      isHiddenColumn,
      mergedCells,
      getValue: getValueText,
      mouseDownInterceptor: handleMouseDownSelection,
      mouseMoveInterceptor: handleMouseMoveSelection,
      canSelectionSpanMergedCells: (start, end) => {
        const { rowCount, columnCount } = internalRefs.current;
        if (
          start.rowIndex === selectionTopBound &&
          end.rowIndex === rowCount - 1
        ) {
          return false;
        }
        if (
          start.columnIndex === selectionLeftBound &&
          end.columnIndex === columnCount - 1
        ) {
          return false;
        }
        return true;
      },
    });

    /**
     * Copy paste
     */
    const { copy, paste, cut } = useCopyPaste({
      gridRef,
      selections,
      activeCell,
      getValue: getValueText,
      onPaste,
      onCut,
    });

    const handleChangeFilter = useCallback(
      (filterIndex: number, columnIndex: number, filter: FilterDefinition) => {
        onChangeFilter?.(filterIndex, columnIndex, filter);
        hideFilter();
      },
      []
    );

    const headerSelections = useMemo(() => {
      const sel: Record<string, Record<string, boolean>> = {
        rows: {},
        cols: {},
      };
      for (let i = 0; i < selections.length; i++) {
        const { bounds } = selections[i];
        const isFullColSelected =
          bounds.top === selectionTopBound && bounds.bottom === rowCount - 1;
        const isFullRowSelected =
          bounds.left === selectionLeftBound &&
          bounds.right === columnCount - 1;
        if (isFullColSelected) {
          for (let j = bounds.left; j <= bounds.right; j++) {
            sel.cols[j] = true;
          }
        }
        if (isFullRowSelected) {
          for (let j = bounds.top; j <= bounds.bottom; j++) {
            sel.rows[j] = true;
          }
        }
      }
      return sel;
    }, [
      selections,
      rowCount,
      selectionTopBound,
      selectionLeftBound,
      columnCount,
    ]);

    const isHeaderSelected = useCallback(
      (index: number, type: string) => {
        return headerSelections[type][index];
      },
      [headerSelections]
    );

    /**
     * Filter
     */
    const { filterComponent, showFilter, hideFilter } = useFilter({
      frozenRows,
      frozenColumns,
      gridRef,
      getFilterComponent: () => {
        return (props) => {
          return (
            <FilterComponent
              {...props}
              onChange={handleChangeFilter}
              onCancel={hideFilter}
            />
          );
        };
      },
      getValue: getValueText,
    });

    const handleFilterClick = useCallback(
      (cell: CellInterface) => {
        const filterIndex = filterViews.findIndex(
          (views) => views.bounds.top === cell.rowIndex
        );
        const filterView = filterViews[filterIndex];
        const currentFilter = filterView?.filters?.[cell.columnIndex];
        if (!filterView) return;
        showFilter(cell, filterIndex, filterView, currentFilter);
      },
      [filterViews, frozenRows, frozenColumns]
    );

    /**
     * Adjusts a column
     */
    const handleAdjustColumn = useCallback(
      (columnIndex: number) => {
        const width =
          getColumnWidth(columnIndex) +
          (columnHasFilter(columnIndex) ? FILTER_ICON_DIM : 0);
        onResize?.(AXIS.X, columnIndex, width);
      },
      [cells, hiddenRows, frozenRows, hiddenColumns, scale]
    );

    /**
     * Check if selections are in
     */
    useEffect(() => {
      /* Row and column headers */
      if (activeCell?.rowIndex === 0 || activeCell?.columnIndex === 0) {
        return;
      }
      onActiveCellChange?.(activeCell, getValue(activeCell)?.text);
    }, [activeCell]);

    /**
     * Save it back to sheet
     */
    useEffect(() => {
      /* Batch this cos of debounce */
      onSheetChangeRef.current?.({ activeCell, selections });

      /* Callback */
      onSelectionChange?.(activeCell, selections);
    }, [selections, activeCell]);

    /**
     * If grid changes, lets restore the state
     */
    useEffect(() => {
      if (scrollState) {
        isTouchDevice
          ? scrollToTouch(scrollState)
          : gridRef.current?.scrollTo(scrollState);
      } else {
        isTouchDevice ? scrollToTopTouch() : gridRef.current?.scrollToTop();
      }
      /* Reset last measures cell */
      gridRef.current?.resetAfterIndices(
        {
          rowIndex: 0,
          columnIndex: 0,
        },
        false
      );
      setActiveCell(initialActiveCell, false);
      setSelections(initialSelections);
      /* Hide filter */
      hideFilter();
      /* Hide editor */
      hideEditor();
    }, [selectedSheet]);

    const handleSubmit = useCallback(
      (
        value: React.ReactText,
        cell: CellInterface,
        nextActiveCell?: CellInterface | null
      ) => {
        onChange?.(value, cell);

        /* Focus on next active cell */
        if (nextActiveCell) setActiveCell(nextActiveCell, true);
      },
      [selectedSheet, cells]
    );

    const { tooltipComponent, ...tooltipProps } = useTooltip({
      getTooltip: (cell) => {
        const cellConfig = getValue(cell);
        const isValid = cellConfig?.valid ?? true;
        const datatype = cellConfig?.datatype;
        const hasError = !!cellConfig?.error;
        const hasTooltip = cellConfig?.tooltip;
        const isHyperLink = datatype === "hyperlink";
        const showTooltip =
          isValid === false || hasError === true || isHyperLink || hasTooltip;
        if (!showTooltip) return null;
        let content: string | undefined;
        if (isValid === false) {
          const validation = cellConfig?.dataValidation;
          content = validation?.prompt;
        }
        if (hasError) {
          content = cellConfig?.error;
        }
        if (hasTooltip) {
          content = cellConfig?.tooltip;
        }
        const variant =
          isValid === false ? "invalid" : hasError ? "error" : void 0;
        const position = isHyperLink ? "bottom" : "right";
        return (props: DefaultTooltipProps) => {
          return (
            <Tooltip
              {...props}
              {...cellConfig}
              position={position}
              content={content}
              variant={variant}
            />
          );
        };
      },
      gridRef,
    });

    /**
     * Editable
     */
    const {
      editorComponent,
      isEditInProgress,
      nextFocusableCell,
      makeEditable,
      setValue,
      hideEditor,
      submitEditor,
      cancelEditor,
      editingCell,
      showEditor,
      ...editableProps
    } = useEditable({
      getEditor: (cell: CellInterface | null) => {
        const config = getValue(cell);
        const type = getEditorType(config?.dataValidation?.type);
        const options = config?.dataValidation?.formulae;
        return (props: EditorProps) => (
          <CellEditor
            {...props}
            background={config?.fill}
            color={config?.color}
            fontSize={config?.fontSize}
            fontFamily={config?.fontFamily}
            underline={config?.underline}
            horizontalAlign={config?.horizontalAlign}
            scale={scale}
            wrap={config?.wrap}
            editorType={type}
            options={options}
          />
        );
      },
      frozenRows,
      frozenColumns,
      isHiddenRow,
      isHiddenColumn,
      gridRef,
      selections,
      activeCell,
      rowCount,
      columnCount,
      selectionTopBound,
      selectionLeftBound,
      onSubmit: handleSubmit,
      getValue: getValueText,
      onChange: onActiveCellValueChange,
      canEdit: (cell: CellInterface) => {
        if (cell.rowIndex === 0 || cell.columnIndex === 0) return false;
        const isReadOnly = getValue(cell)?.readOnly;
        if (isReadOnly) return false;
        return true;
      },
      onDelete: onDelete,
    });

    /* Ref to store event references for editable and selection. This prevents re-rendering grid */
    const eventRefs = useRef({
      selectionProps: {
        onMouseDown: selectionProps.onMouseDown,
        onKeyDown: selectionProps.onKeyDown,
      },
      editableProps: {
        onMouseDown: editableProps.onMouseDown,
        onKeyDown: editableProps.onKeyDown,
      },
    });

    useEffect(() => {
      eventRefs.current = {
        selectionProps: {
          onMouseDown: selectionProps.onMouseDown,
          onKeyDown: selectionProps.onKeyDown,
        },
        editableProps: {
          onMouseDown: editableProps.onMouseDown,
          onKeyDown: editableProps.onKeyDown,
        },
      };
    });

    const getNextFocusableCell = useCallback(
      (cell: CellInterface, direction: Direction): CellInterface | null => {
        return nextFocusableCell(cell, direction);
      },
      [hiddenRows, hiddenColumns]
    );

    /* Width calculator */
    const columnWidth = useCallback(
      (columnIndex: number) => {
        if (hiddenColumns[columnIndex]) return 0;
        if (columnIndex === 0) return COLUMN_HEADER_WIDTH;
        return columnSizes[columnIndex] ?? minColumnWidth;
      },
      [minColumnWidth, columnSizes, selectedSheet]
    );
    const rowHeight = useCallback(
      (rowIndex: number) => {
        if (hiddenRows[rowIndex]) return 0;
        if (rowIndex === 0) return ROW_HEADER_HEIGHT;
        return rowSizes[rowIndex] ?? minRowHeight;
      },
      [minRowHeight, hiddenRows, rowSizes, selectedSheet]
    );
    const contextWrapper = useCallback(
      (children) => {
        return (
          <ThemeProvider theme={theme}>
            <ColorModeProvider>{children}</ColorModeProvider>
          </ThemeProvider>
        );
      },
      [theme]
    );
    const selectedRowsAndCols: RowColSelection = useMemo(() => {
      const activeCellBounds: SelectionArea[] =
        activeCell && gridRef.current
          ? [{ bounds: gridRef.current?.getCellBounds?.(activeCell) }]
          : [];
      const initial: SelectionArea[] = [];
      return initial.concat(selections, activeCellBounds).reduce(
        (acc, { bounds }) => {
          for (let i = bounds.left; i <= bounds.right; i++) {
            acc.cols.push(i);
          }
          for (let i = bounds.top; i <= bounds.bottom; i++) {
            acc.rows.push(i);
          }
          return acc;
        },
        { rows: [], cols: [] } as RowColSelection
      );
    }, [selections, activeCell, gridRef.current?.getCellBounds]);

    /**
     * Check if a row is selected
     */
    const isRowSelected = useCallback(
      (rowIndex) => {
        if (selectionMode === SELECTION_MODE.CELL) return false;
        return selectionMode === SELECTION_MODE.ROW ||
          selectionMode === SELECTION_MODE.BOTH
          ? activeCell?.rowIndex === rowIndex ||
              selectedRowsAndCols.rows.includes(rowIndex)
          : false;
      },
      [selectedRowsAndCols, activeCell, selectionMode]
    );

    /**
     * Check if a column is selected
     */
    const isColumnSelected = useCallback(
      (columnIndex) => {
        if (selectionMode === SELECTION_MODE.CELL) return false;
        return selectionMode === SELECTION_MODE.COLUMN ||
          selectionMode === SELECTION_MODE.BOTH
          ? activeCell?.columnIndex === columnIndex ||
              selectedRowsAndCols.cols.includes(columnIndex)
          : false;
      },
      [selectedRowsAndCols, activeCell, selectionMode]
    );

    /**
     * Check if a cell should be hidden
     */
    const isHiddenCell = useCallback(
      (rowIndex: number, columnIndex: number) => {
        const rowSelected = isRowSelected(rowIndex);
        const columnSelected = isColumnSelected(columnIndex);
        if (rowSelected || columnSelected) return false;

        const isRowHeader = isCellRowHeader(rowIndex);
        const isColumnHeader = isCellColumnHeader(columnIndex);
        if (isRowHeader || isColumnHeader) return false;
        return cells?.[rowIndex]?.[columnIndex] === void 0;
      },
      [cells, selectionMode, selectedRowsAndCols]
    );

    /**
     * Switch cell to edit mode
     */
    const handleEdit = useCallback((cell: CellInterface) => {
      makeEditable(cell);
    }, []);

    /**
     * Called when checkbox is checked/unchecked
     */
    const handleCheck = useCallback((cell: CellInterface, checked: boolean) => {
      const cellConfig = getValueRef.current(cell);
      const type = cellConfig?.dataValidation?.type;
      const formulae: string[] =
        cellConfig?.dataValidation?.formulae ?? DEFAULT_CHECKBOX_VALUES;
      if (!type) {
        console.error(
          "Type is not specified. Cell Config should contain a dataValidation object",
          cellConfig
        );
        return;
      }
      const text = formulae[checked ? 0 : 1];
      onChange?.(text, cell);
      onActiveCellValueChange?.(text, cell);
    }, []);

    const itemRenderer = useCallback(
      (props: RendererProps) => {
        const { rowIndex, columnIndex } = props;
        const cell = { rowIndex, columnIndex };
        const isRowHeader = isCellRowHeader(rowIndex);
        const isColumnHeader = isCellColumnHeader(columnIndex);
        if (isRowHeader || isColumnHeader) {
          const isHeaderActive =
            isRowHeader || isColumnHeader
              ? isRowHeader
                ? selectedRowsAndCols.cols.includes(columnIndex)
                : selectedRowsAndCols.rows.includes(rowIndex)
              : false;
          const isFilteredColumn = columnHasFilter(columnIndex);
          const isFilteredRow = rowHasFilter(rowIndex);
          const isSelected = isHeaderSelected(
            isRowHeader ? columnIndex : rowIndex,
            isRowHeader ? "cols" : "rows"
          );
          return (
            <HeaderCellRenderer
              {...props}
              isLightMode={isLightMode}
              isActive={isHeaderActive}
              onResize={onResize}
              onAdjustColumn={handleAdjustColumn}
              theme={theme}
              scale={scale}
              isFiltered={isFilteredColumn || isFilteredRow}
              isSelected={isSelected}
            />
          );
        }
        /* Row, cell column selection modes */
        const rowSelected = isRowSelected(rowIndex);
        const columnSelected = isColumnSelected(columnIndex);
        const isSelected = rowSelected || columnSelected;
        const filterIndex =
          filterHeaderCells[cellIdentifier(rowIndex, columnIndex)];
        const showFilter = filterIndex !== void 0;
        const cellConfig = getValue(cell);
        const isFilterActive =
          filterIndex === void 0
            ? false
            : !!filterViews?.[filterIndex]?.filters?.[columnIndex];
        return (
          <CellRenderer
            {...props}
            {...cellConfig}
            cellConfig={cellConfig}
            isLightMode={isLightMode}
            formatter={formatter}
            showStrokeOnFill={showGridLines}
            isSelected={isSelected}
            showFilter={showFilter}
            isFilterActive={isFilterActive}
            onFilterClick={handleFilterClick}
            scale={scale}
            onEdit={handleEdit}
            onCheck={handleCheck}
          />
        );
      },
      [
        cells,
        selectedRowsAndCols,
        activeCell,
        showGridLines,
        selectionMode,
        isLightMode,
        theme,
        scale,
        filterHeaderCells,
      ]
    );

    const overlayRenderer = useCallback(
      (props: RendererProps) => {
        const { rowIndex, columnIndex } = props;
        const cell = { rowIndex, columnIndex };
        const cellConfig = getValue(cell);
        return cellOverlayRenderer({
          ...props,
          ...cellConfig,
        });
      },
      [cells, selectedRowsAndCols, activeCell, hiddenRows, hiddenColumns]
    );

    const handleScroll = useCallback(
      (scrollState: ScrollCoords) => {
        editableProps.onScroll?.(scrollState);
        // Save scroll state in sheet
        debounceScroll.current?.(scrollState);
      },
      [selectedSheet]
    );

    /**
     * Show context menu
     */
    const showContextMenu = useCallback(
      (e: React.MouseEvent<HTMLDivElement>) => {
        e.preventDefault();
        const left = e.clientX;
        const top = e.clientY;
        const pos = gridRef.current?.getRelativePositionFromOffset(left, top);
        if (!pos) return;
        const { x, y } = pos;
        setContextMenuProps({
          left: x,
          top: y,
        });
      },
      []
    );

    const handleMouseDown = useCallback(
      (e: React.MouseEvent<HTMLDivElement>) => {
        eventRefs.current.selectionProps.onMouseDown(e);
        eventRefs.current.editableProps.onMouseDown(e);
        hideContextMenu();
        hideFilter();
      },
      [cells, selectedSheet, activeCell, selections]
    );

    const handleKeyDown = useCallback(
      (e: React.KeyboardEvent<HTMLDivElement>) => {
        eventRefs.current.selectionProps.onKeyDown(e);
        eventRefs.current.editableProps.onKeyDown(e);
        onKeyDown?.(e, activeCell, selections);
      },
      [cells, selectedSheet, activeCell, selections]
    );

    /**
     * Hides context menu
     */
    const hideContextMenu = useCallback(() => {
      setContextMenuProps(null);
      gridRef.current?.focus();
    }, []);
    const fillhandleBorderColor = isLightMode ? "white" : DARK_MODE_COLOR_LIGHT;
    const gridLineColor = isLightMode
      ? CELL_BORDER_COLOR
      : theme?.colors.gray[600];
    const shadowStroke = isLightMode
      ? HEADER_BORDER_COLOR
      : theme?.colors.gray[600];
    const shadowSettings = useMemo(() => {
      return {
        stroke: shadowStroke,
      };
    }, [shadowStroke]);

    return (
      <GridWrapper>
        <Grid
          snap={snap}
          scale={scale}
          isHiddenRow={isHiddenRow}
          isHiddenCell={isHiddenCell}
          enableCellOverlay
          shadowSettings={shadowSettings}
          showFrozenShadow
          gridLineColor={gridLineColor}
          showGridLines={showGridLines}
          ref={gridRef}
          rowCount={rowCount}
          columnCount={columnCount}
          columnWidth={columnWidth}
          rowHeight={rowHeight}
          width={width}
          height={height}
          itemRenderer={itemRenderer}
          overlayRenderer={overlayRenderer}
          wrapper={contextWrapper}
          selections={selections}
          activeCell={activeCell}
          fillhandleBorderColor={fillhandleBorderColor}
          showFillHandle={!isEditInProgress}
          mergedCells={mergedCells}
          borderStyles={borderStyles}
          frozenRows={frozenRows}
          frozenColumns={frozenColumns}
          {...selectionProps}
          {...editableProps}
          onMouseDown={handleMouseDown}
          onKeyDown={handleKeyDown}
          onScroll={handleScroll}
          onContextMenu={showContextMenu}
          onViewChange={onViewChange}
          {...tooltipProps}
        />
        {editorComponent}
        {ContextMenu !== void 0 && contextMenuProps && (
          <ContextMenu
            onRequestClose={hideContextMenu}
            activeCell={activeCell}
            selections={selections}
            onCopy={copy}
            onPaste={paste}
            onCut={cut}
            onInsertRow={onInsertRow}
            onInsertColumn={onInsertColumn}
            onDeleteColumn={onDeleteColumn}
            onDeleteRow={onDeleteRow}
            {...contextMenuProps}
          />
        )}
        {filterComponent}
        {tooltipComponent}
      </GridWrapper>
    );
  })
);

export default SheetGrid;
