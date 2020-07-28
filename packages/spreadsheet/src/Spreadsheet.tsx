import React, {
  useState,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  forwardRef,
  useImperativeHandle,
  memo,
  useReducer
} from "react";
import Toolbar from "./Toolbar";
import Formulabar from "./Formulabar";
import Workbook from "./Workbook";
import { theme, ThemeProvider, ColorModeProvider, Flex } from "@chakra-ui/core";
import { Global, css } from "@emotion/core";
import {
  CellInterface,
  SelectionArea,
  ScrollCoords,
  AreaProps,
  FilterView,
  FilterDefinition,
  useUndo,
  SelectionPolicy
} from "@rowsncolumns/grid";
import {
  createNewSheet,
  uuid,
  DEFAULT_COLUMN_WIDTH,
  DEFAULT_ROW_HEIGHT,
  SYSTEM_FONT,
  format as defaultFormat,
  FONT_FAMILIES,
  detectDataType,
  getMinMax,
  castToString
} from "./constants";
import {
  FORMATTING_TYPE,
  CellFormatting,
  AXIS,
  BORDER_VARIANT,
  BORDER_STYLE,
  Formatter,
  SELECTION_MODE
} from "./types";
import { WorkbookGridRef } from "./Grid/Grid";
import { KeyCodes, Direction } from "@rowsncolumns/grid/dist/types";
import invariant from "tiny-invariant";
import { ThemeType } from "./styled";
import Editor, { CustomEditorProps } from "./Editor/Editor";
import StatusBarComponent from "./StatusBar";
import { StatusBarProps } from "./StatusBar/StatusBar";
import useFonts from "./hooks/useFonts";
import {
  createStateReducer,
  ACTION_TYPE,
  StateInterface,
  ActionTypes
} from "./state";
import { Patch } from "immer";
import { ContextMenuComponentProps } from "./ContextMenu/ContextMenu";
import ContextMenuComponent from "./ContextMenu";
import TooltipComponent, { TooltipProps } from "./Tooltip";
import validate, { ValidationResponse } from "./validation";

export interface SpreadSheetProps {
  /**
   * Minimum column width of the grid
   */
  minColumnWidth?: number;
  /**
   * Minimum row height of the grid
   */
  minRowHeight?: number;
  /**
   * Customize cell rendering
   */
  CellRenderer?: React.ReactType;
  /**
   * Custom header cell
   */
  HeaderCellRenderer?: React.ReactType;
  /**
   * Array of sheets to render
   */
  sheets?: Sheet[];
  /**
   * Uncontrolled sheets
   */
  initialSheets?: Sheet[];
  /**
   * Active  sheet on the workbook
   */
  activeSheet?: string;
  /**
   * Callback fired when cells are modified
   */
  onChangeCell?: (
    id: SheetID,
    value: React.ReactText,
    cell: CellInterface
  ) => void;
  /**
   * Get the new selected sheet
   */
  onChangeSelectedSheet?: (id: SheetID) => void;
  /**
   * Listen to changes to all the sheets
   */
  onChange?: (sheets: Sheet[]) => void;
  /**
   * Show formula bar
   */
  showFormulabar?: boolean;
  /**
   * Show hide toolbar
   */
  showToolbar?: boolean;
  /**
   * Conditionally format cell text
   */
  formatter?: Formatter;
  /**
   * Enabled or disable dark mode
   */
  enableDarkMode?: true;
  /**
   * Font family
   */
  fontFamily?: string;
  /**
   * Min Height of the grid
   */
  minHeight?: number;
  /**
   * Custom Cell Editor
   */
  CellEditor?: React.ReactType<CustomEditorProps>;
  /**
   * Allow user to customize single, multiple or range selection
   */
  selectionPolicy?: SelectionPolicy;
  /**
   * Callback when active cell changes
   */
  onActiveCellChange?: (
    /* Sheet id */
    id: SheetID,
    /* Cell coords */
    cell: CellInterface | null,
    /* Value of the active cell */
    value?: React.ReactText
  ) => void;

  /**
   * Callback When active cell values changes
   */
  onActiveCellValueChange?: (
    /* Sheet id */
    id: SheetID,
    /* Cell coords  */
    cell: CellInterface | null,
    /* Value of the active cell */
    value?: React.ReactText
  ) => void;
  /**
   * Callback fired when selection changes
   */
  onSelectionChange?: (
    id: SheetID,
    activeCell: CellInterface | null,
    selections: SelectionArea[]
  ) => void;
  /**
   * Select mode
   */
  selectionMode?: SELECTION_MODE;
  /**
   * Show or hide tab strip
   */
  showTabStrip?: boolean;
  /**
   * Make tab editable
   */
  isTabEditable?: boolean;
  /**
   * Allow user to add new sheet
   */
  allowNewSheet?: boolean;
  /**
   * Show or hide status bar
   */
  showStatusBar?: boolean;
  /**
   * Status bar component
   */
  StatusBar?: React.ReactType<StatusBarProps>;
  /**
   * Context menu component
   */
  ContextMenu?: React.ReactType<ContextMenuComponentProps>;
  /**
   * Tooltip component
   */
  Tooltip?: React.ReactType<TooltipProps>;
  /**
   * Scale
   */
  initialScale?: number;
  /**
   * When scale changes
   */
  onScaleChange?: (scale: number) => void;
  /**
   * Web font loader config
   */
  fontLoaderConfig?: WebFont.Config;
  /**
   * Visible font families
   */
  fontList?: string[];
  /**
   * Snap to row and column as user scrolls
   */
  snap?: boolean;
  /**
   * Add your own state interface
   */
  stateReducer?: (state: StateInterface, action: ActionTypes) => StateInterface;
  /**
   * Custom onvalidator
   */
  onValidate?: (
    value: React.ReactText,
    id: SheetID,
    cell: CellInterface,
    cellConfig: CellConfig | undefined
  ) => Promise<ValidationResponse>;
  /**
   * By default, all keydown listeners are bound to the grid.
   * If you want to bind listeners to `document` for events such as undo/redo,
   * Toggle this to true
   */
  enableGlobalKeyHandlers?: boolean;
  /**
   * Called when the grid is initialized,
   * so that formula module can add dependencies in the graph
   */
  onInitialize?: (changes: CellsBySheet, getCellConfig: CellConfigGetter | undefined) => Promise<CellsBySheet> | undefined
  /**
   * 
   */
  onCalculate?: (value: React.ReactText, id: SheetID, cell: CellInterface, getCellConfig?: CellConfigGetter) => Promise<CellsBySheet>
  // TODO
  // onMouseOver?: (event: React.MouseEvent<HTMLDivElement>, cell: CellInterface) => void;
  // onMouseDown?: (event: React.MouseEvent<HTMLDivElement>, cell: CellInterface) => void;
  // onMouseUp?: (event: React.MouseEvent<HTMLDivElement>, cell: CellInterface) => void;
  // onClick?: (event: React.MouseEvent<HTMLDivElement>, cell: CellInterface) => void;
}

export type CellConfigGetter = (id: SheetID, cell: CellInterface | null) => CellConfig | undefined

export interface Sheet {
  id: SheetID;
  name: string;
  cells: Cells;
  activeCell: CellInterface | null;
  selections: SelectionArea[];
  scrollState?: ScrollCoords;
  columnSizes?: SizeType;
  rowSizes?: SizeType;
  mergedCells?: AreaProps[];
  frozenRows?: number;
  frozenColumns?: number;
  hiddenRows?: number[];
  hiddenColumns?: number[];
  showGridLines?: boolean;
  filterViews?: FilterView[];
  rowCount?: number;
  columnCount?: number;
  locked?: boolean;
  hidden?: boolean;
  tabColor?: string;
}

export type SheetID = React.ReactText;
export type CellsBySheet = Record<string, Cells>

export type SizeType = {
  [key: number]: number;
};

export type Cells = Record<string, Cell>;
export type Cell = Record<string, CellConfig>;
export interface CellConfig extends CellFormatting {
  /**
   * Text that will be displayed in the cell.
   * For formulas, result will be displayed instead
   */
  text?: string | number;
  /**
   * Add tooltip
   */
  tooltip?: string;
  /**
   * Result from formula calculation
   */
  result?: string | number | boolean | Date;
  /**
   * Formula errors
   */
  error?: string;
  /**
   * Validation errors
   */
  valid?: boolean;
}

const defaultActiveSheet = uuid();
export const defaultSheets: Sheet[] = [
  {
    id: defaultActiveSheet,
    name: "Sheet1",
    rowCount: 1000,
    columnCount: 26,
    frozenColumns: 0,
    frozenRows: 0,
    activeCell: {
      rowIndex: 1,
      columnIndex: 1
    },
    mergedCells: [],
    selections: [],
    cells: {},
    scrollState: { scrollTop: 0, scrollLeft: 0 },
    filterViews: []
  }
];

export type RefAttributeSheetGrid = {
  ref?: React.Ref<SheetGridRef>;
};

export type SheetGridRef = {
  grid: WorkbookGridRef | null;
};

export interface PatchInterface {
  patches: Patch;
  inversePatches: Patch;
}

/**
 * Spreadsheet component
 * TODO
 * 1. Undo/redo
 * @param props
 */
const Spreadsheet: React.FC<SpreadSheetProps & RefAttributeSheetGrid> = memo(
  forwardRef((props, forwardedRef) => {
    const {
      sheets: initialSheets = defaultSheets,
      showFormulabar = true,
      minColumnWidth = DEFAULT_COLUMN_WIDTH,
      minRowHeight = DEFAULT_ROW_HEIGHT,
      CellRenderer,
      HeaderCellRenderer,
      activeSheet,
      onChangeSelectedSheet,
      onChange,
      onChangeCell,
      showToolbar = true,
      formatter = defaultFormat,
      enableDarkMode = true,
      fontFamily = SYSTEM_FONT,
      minHeight = 400,
      CellEditor = Editor,
      onActiveCellChange,
      onActiveCellValueChange,
      onSelectionChange,
      selectionMode,
      showTabStrip = true,
      isTabEditable = true,
      allowNewSheet = true,
      showStatusBar = true,
      StatusBar = StatusBarComponent,
      ContextMenu = ContextMenuComponent,
      Tooltip = TooltipComponent,
      initialScale = 1,
      onScaleChange,
      fontLoaderConfig,
      fontList = FONT_FAMILIES,
      selectionPolicy,
      snap = false,
      stateReducer,
      onValidate = validate,
      enableGlobalKeyHandlers = false,
      onInitialize,
      onCalculate
    } = props;

    /* Last active cells: for undo, redo */
    const lastActiveCellRef = useRef<CellInterface | null | undefined>(null);
    const lastSelectionsRef = useRef<SelectionArea[] | null | undefined>([]);

    /**
     * Some grid side-effects during undo/redo
     */
    const beforeUndoRedo = useCallback((patches: Patch[]) => {
      const hasFilterViews = patches.some(item =>
        item.path.includes("filterViews")
      );
      if (hasFilterViews) {
        currentGrid.current?.resetAfterIndices?.({ rowIndex: 0 }, false);
      }
    }, []);

    /**
     * Undo hook
     */
    const {
      add: addUndoPatch,
      canRedo,
      canUndo,
      undo,
      redo,
      onKeyDown: onUndoKeyDown
    } = useUndo<Patch[]>({
      enableGlobalKeyHandlers,
      onUndo: patches => {
        /* Side-effects */
        beforeUndoRedo(patches);

        dispatch({
          type: ACTION_TYPE.APPLY_PATCHES,
          patches,
          undoable: false
        });

        if (lastActiveCellRef.current) {
          currentGrid.current?.setActiveCell(lastActiveCellRef.current);
        }
        if (lastSelectionsRef.current) {
          currentGrid.current?.setSelections(lastSelectionsRef.current);
        }
        /* Focus on the grid */
        if (enableGlobalKeyHandlers) currentGrid.current?.focus();
      },
      onRedo: patches => {
        /* Side-effects */
        beforeUndoRedo(patches);

        dispatch({
          type: ACTION_TYPE.APPLY_PATCHES,
          patches,
          undoable: false
        });

        const activeCellPatch = patches.find((item: Patch) =>
          item.path.includes("currentActiveCell")
        );
        const selectionsPatch = patches.find((item: Patch) =>
          item.path.includes("currentSelections")
        );
        if (activeCellPatch) {
          currentGrid.current?.setActiveCell(activeCellPatch.value);
        }
        if (selectionsPatch) {
          currentGrid.current?.setSelections(selectionsPatch.value);
        }

        /* Focus on the grid */
        if (enableGlobalKeyHandlers) currentGrid.current?.focus();
      }
    });

    /**
     * Get cell bounds
     */
    const getCellBounds = useCallback((cell: CellInterface | null) => {
      if (!cell) return undefined;
      return currentGrid.current?.getCellBounds?.(cell);
    }, []);

    /**
     * State reducer
     */
    const [state, dispatch] = useReducer(
      useCallback(
        createStateReducer({ addUndoPatch, getCellBounds, stateReducer }),
        []
      ),
      {
        sheets: initialSheets,
        selectedSheet:
          activeSheet === void 0
            ? initialSheets.length
              ? initialSheets[0].id
              : null
            : activeSheet
      }
    );

    const {
      selectedSheet,
      sheets,
      currentActiveCell,
      currentSelections
    } = state;
    const [scale, setScale] = useState(initialScale);
    const selectedSheetRef = useRef(selectedSheet);
    const currentGrid = useRef<WorkbookGridRef>(null);
    const [formulaInput, setFormulaInput] = useState("");

    /* Last */
    useEffect(() => {
      lastActiveCellRef.current = currentActiveCell;
      lastSelectionsRef.current = currentSelections;
    }, [currentActiveCell, currentSelections]);

    /* Selected sheet */
    const setSelectedSheet = useCallback(
      (id: React.ReactText) => {
        if (id === selectedSheet) return;
        dispatch({
          type: ACTION_TYPE.SELECT_SHEET,
          id
        });
      },
      [selectedSheet]
    );

    invariant(
      selectedSheet !== null,
      "Exception, selectedSheet is empty, Please specify a selected sheet using `selectedSheet` prop"
    );

    /* Fonts */
    const { isFontActive } = useFonts(fontLoaderConfig);

    /* Callback fired when fonts are loaded */
    useEffect(() => {
      if (isFontActive) {
        currentGrid.current?.resetAfterIndices?.({
          rowIndex: 0,
          columnIndex: 0
        });
      }
    }, [isFontActive]);

    /* Callback fired when scale changes */
    useEffect(() => {
      onScaleChange?.(scale);
    }, [scale]);

    useImperativeHandle(
      forwardedRef,
      () => {
        return {
          grid: currentGrid.current
        };
      },
      []
    );

    /* Callback when sheets is changed */
    useEffect(() => {
      onChange?.(sheets);
    }, [sheets]);

    /* Change selected sheet */
    useEffect(() => {
      onChangeSelectedSheet?.(selectedSheet);
      selectedSheetRef.current = selectedSheet;
    }, [selectedSheet]);

    /* Listen to sheet change */
    useEffect(() => {
      /* If its the same sheets - Skip */
      if (sheets === initialSheets) {
        return;
      }

      dispatch({
        type: ACTION_TYPE.REPLACE_SHEETS,
        sheets: initialSheets
      });
    }, [initialSheets]);    

    /**
     * Handle add new sheet
     */
    const handleNewSheet = useCallback(() => {
      const count = sheets.length;
      const newSheet = createNewSheet({ count: count + 1 });
      dispatch({
        type: ACTION_TYPE.NEW_SHEET,
        sheet: newSheet
      });

      /* Focus on the new grid */
      currentGrid.current?.focus();
    }, [sheets, selectedSheet]);

    const sheetsById = useMemo(() => {
      const initial: Record<string, Sheet> = {};
      return sheets.reduce((acc, sheet) => {
        acc[sheet.id] = sheet;
        return acc;
      }, initial);
    }, [sheets]);

    /* Current sheet */
    const currentSheet = sheetsById[selectedSheet];

    /**
     * Get cell config
     */
    const getCellConfig = useCallback(
      (id: SheetID, cell: CellInterface | null): CellConfig | undefined => {
        if (!cell) return void 0
        return sheetsById?.[id].cells?.[cell.rowIndex]?.[cell.columnIndex];
      },
      [sheetsById]
    );

    /* Add it to ref to prevent closures */
    const getCellConfigRef = useRef<CellConfigGetter>()
    useEffect(() => {
      getCellConfigRef.current = getCellConfig
    }, [ getCellConfig])

    const getSheet = useCallback((id: SheetID) => {
      return sheetsById?.[id]
    }, [ sheetsById ])

    useEffect(() => {
      const initial: CellsBySheet =  {}
      const changes = sheets.reduce((acc, sheet) => {
        // @ts-nocheck
        acc[sheet.id] = sheet.cells
        return acc
      }, initial)
      
      async function triggerBatchCalculation (changes: CellsBySheet) {
        const values = await onInitialize?.(changes, getCellConfigRef.current)
        if (values !== void 0) {
          dispatch({
            type: ACTION_TYPE.UPDATE_CELLS,
            changes: values,
          })
        }
      }
      /* Trigger batch calculation */
      triggerBatchCalculation(changes)
      
    }, [])
    
    /**
     * Get max rows in a sheet
     */
    const getMinMaxRows = useCallback((sheet: SheetID) => {
      return getMinMax(sheetsById[sheet]?.cells)  
    }, [getCellConfig])

    /**
     * Get max columns in a sheet row
     */
    const getMinMaxColumns = useCallback((sheet: SheetID, rowIndex: number) => {
      return getMinMax(sheetsById[sheet].cells?.[rowIndex])  
    }, [getCellConfig])

    /**
     * Active cell + Active cell config.
     * Used in toolbars
     */
    const [ activeCellConfig, activeCell ] = useMemo(() => {
      const sheet = getSheet(selectedSheet)
      return [
        getCellConfig(selectedSheet, sheet.activeCell),
        sheet.activeCell
      ]
    }, [ getSheet, selectedSheet, getCellConfig ])

    /**
     * Cell changes on user input
     * General purpos changes
     */
    const handleChange = useCallback(
      async (id: SheetID, value: React.ReactText, cell: CellInterface) => {
        const config = getCellConfig(id, cell);
        const datatype = detectDataType(value);

        dispatch({
          type: ACTION_TYPE.CHANGE_SHEET_CELL,
          value,
          cell,
          id,
          datatype
        });

        onChangeCell?.(id, value, cell);

        /* Validate */
        requestAnimationFrame(async () => {
          const validationResponse = await onValidate(value, id, cell, config);

          /* If validations service fails, lets not update the store */
          if (validationResponse !== void 0) {
            /**
             * Extract valid: boolean response and message
             */
            const { valid, message } = validationResponse as ValidationResponse;

            /**
             * Update the state
             */
            dispatch({
              type: ACTION_TYPE.VALIDATION_SUCCESS,
              cell,
              id,
              valid,
              prompt: message
            });
          }

          const changes = await onCalculate?.(value, id, cell, getCellConfigRef.current)

          if (changes !== void 0) {
            dispatch({
              type: ACTION_TYPE.UPDATE_CELLS,
              changes,
            })
          }

        });
      },
      [getCellConfig]
    );

    const handleSheetAttributesChange = useCallback(
      (
        id: SheetID,
        {
          activeCell,
          selections
        }: { activeCell: CellInterface | null; selections: SelectionArea[] }
      ) => {
        dispatch({
          type: ACTION_TYPE.SHEET_SELECTION_CHANGE,
          id,
          activeCell,
          selections,
          undoable: false
        });
      },
      []
    );

    /**
     * Handle sheet name
     */
    const handleChangeSheetName = useCallback((id: SheetID, name: string) => {
      dispatch({
        type: ACTION_TYPE.CHANGE_SHEET_NAME,
        id,
        name
      });
    }, []);

    const handleDeleteSheet = useCallback(
      (id: SheetID) => {
        if (sheets.length === 1) return;

        dispatch({
          type: ACTION_TYPE.DELETE_SHEET,
          id
        });

        /* Focus on the new grid */
        currentGrid.current?.focus();
      },
      [sheets, selectedSheet]
    );

    const handleDuplicateSheet = useCallback(
      (id: SheetID) => {
        const newSheetId = uuid();
        const index = sheets.findIndex(sheet => sheet.id === id);
        if (index === -1) return;
        const newSheet = {
          ...sheets[index],
          locked: false,
          id: newSheetId,
          name: `Copy of ${currentSheet.name}`
        };

        dispatch({
          type: ACTION_TYPE.NEW_SHEET,
          sheet: newSheet,
          index
        });
      },
      [sheets, selectedSheet]
    );

    /**
     * Change formatting to auto
     */
    const handleFormattingChangeAuto = useCallback(() => {
      dispatch({
        type: ACTION_TYPE.FORMATTING_CHANGE_AUTO,
        id: selectedSheet
      });
    }, [selectedSheet]);

    /**
     * Change formatting to plain
     */
    const handleFormattingChangePlain = useCallback(() => {
      dispatch({
        type: ACTION_TYPE.FORMATTING_CHANGE_PLAIN,
        id: selectedSheet
      });
    }, [selectedSheet]);

    /**
     * When cell or selection formatting change
     */
    const handleFormattingChange = useCallback(
      (key, value) => {
        dispatch({
          type: ACTION_TYPE.FORMATTING_CHANGE,
          id: selectedSheet,
          key,
          value
        });
      },
      [selectedSheet]
    );

    const handleActiveCellChange = useCallback(
      (id: SheetID, cell: CellInterface | null, value) => {
        if (!cell) return;
        setFormulaInput(value || "");
        onActiveCellChange?.(id, cell, value);
      },
      []
    );

    const handleActiveCellValueChange = useCallback(
      (id: SheetID, activeCell: CellInterface | null, value) => {
        setFormulaInput(value);
        onActiveCellValueChange?.(id, activeCell, value);
      },
      []
    );

    /**
     * Formula bar focus event
     */
    const handleFormulabarFocus = useCallback(
      (e: React.FocusEvent<HTMLInputElement>) => {
        const input = e.target;
        if (activeCell) {
          currentGrid.current?.makeEditable(activeCell, input.value, false);
          requestAnimationFrame(() => input?.focus());
        }
      },
      [activeCell]
    );

    /**
     * When formula input changes
     */
    const handleFormulabarChange = useCallback(
      (e: React.ChangeEvent<HTMLInputElement>) => {
        if (!activeCell) return;
        const value = e.target.value;
        setFormulaInput(value);
        /* Row and column headers */
        if (activeCell?.rowIndex === 0 || activeCell?.columnIndex === 0) {
          return;
        }
        currentGrid.current?.setEditorValue(value, activeCell);
      },
      [activeCell, selectedSheet]
    );

    /**
     * Imperatively submits the editor
     * @param value
     * @param activeCell
     */
    const submitEditor = (
      value: string,
      activeCell: CellInterface,
      direction: Direction = Direction.Down
    ) => {
      const nextActiveCell = currentGrid.current?.getNextFocusableCell(
        activeCell,
        direction
      );
      currentGrid.current?.submitEditor(value, activeCell, nextActiveCell);
    };
    /**
     * When user presses Enter on formula input
     */
    const handleFormulabarKeydown = useCallback(
      (e: React.KeyboardEvent<HTMLInputElement>) => {
        if (!activeCell) return;
        /* Row and column headers */
        if (activeCell?.rowIndex === 0 || activeCell?.columnIndex === 0) {
          return;
        }

        if (e.which === KeyCodes.Enter) {
          submitEditor(formulaInput, activeCell);
        }
        if (e.which === KeyCodes.Escape) {
          currentGrid.current?.cancelEditor();
          setFormulaInput(activeCellConfig?.text?.toString() || "");
        }
        if (e.which === KeyCodes.Tab) {
          submitEditor(formulaInput, activeCell, Direction.Right);
          e.preventDefault();
        }
      },
      [activeCell, formulaInput, activeCellConfig]
    );

    /**
     * Handle fill
     */
    const handleFill = useCallback(
      (
        id: SheetID,
        activeCell: CellInterface,
        fillSelection: SelectionArea | null,
        selections: SelectionArea[]
      ) => {
        if (!fillSelection) return;

        dispatch({
          type: ACTION_TYPE.UPDATE_FILL,
          id,
          activeCell,
          fillSelection,
          selections
        });

        /* Focus on the new grid */
        currentGrid.current?.focus();
      },
      [sheets]
    );

    /**
     * Delete cell values
     */
    const handleDelete = useCallback(
      (id: SheetID, activeCell: CellInterface, selections: SelectionArea[]) => {
        dispatch({
          type: ACTION_TYPE.DELETE_CELLS,
          id,
          activeCell,
          selections
        });
        setFormulaInput("");
      },
      []
    );
    
    /**
     * Clear formatting of selected area
     */
    const handleClearFormatting = useCallback(() => {
      dispatch({
        type: ACTION_TYPE.CLEAR_FORMATTING,
        id: selectedSheet
      });
    }, [selectedSheet]);

    /**
     * Trigger row or column resize
     */
    const handleResize = useCallback(
      (id: SheetID, axis: AXIS, index: number, dimension: number) => {
        dispatch({
          type: ACTION_TYPE.RESIZE,
          id,
          index,
          dimension,
          axis,
          undoable: false
        });

        axis === AXIS.X
          ? currentGrid.current?.resizeColumns?.([index])
          : currentGrid.current?.resizeRows?.([index]);
      },
      []
    );

    /**
     * Handle toggle cell merges
     */
    const handleMergeCells = useCallback(() => {
      dispatch({
        type: ACTION_TYPE.MERGE_CELLS,
        id: selectedSheet
      });
    }, [selectedSheet]);

    const handleFrozenRowChange = useCallback(
      count => {
        dispatch({
          type: ACTION_TYPE.FROZEN_ROW_CHANGE,
          id: selectedSheet,
          count
        });
      },
      [selectedSheet]
    );

    const handleFrozenColumnChange = useCallback(
      count => {
        dispatch({
          type: ACTION_TYPE.FROZEN_COLUMN_CHANGE,
          id: selectedSheet,
          count
        });
      },
      [selectedSheet]
    );

    const handleBorderChange = useCallback(
      (
        color: string | undefined,
        borderStyle: BORDER_STYLE,
        variant?: BORDER_VARIANT
      ) => {
        dispatch({
          type: ACTION_TYPE.SET_BORDER,
          id: selectedSheet,
          color,
          borderStyle,
          variant
        });
      },
      [selectedSheet]
    );

    /**
     * Handle sheet scroll
     */
    const handleScroll = useCallback(
      (id: SheetID, scrollState: ScrollCoords) => {
        dispatch({
          type: ACTION_TYPE.UPDATE_SCROLL,
          id,
          scrollState
        });
      },
      []
    );

    /**.
     * On Paste
     * TODO: Preserve formatting
     */
    const handlePaste = useCallback(
      (
        id: SheetID,
        rows,
        activeCell: CellInterface | null,
        selection?: SelectionArea
      ) => {
        if (!activeCell) return;
        const { rowIndex, columnIndex } = activeCell;
        const endRowIndex = Math.max(rowIndex, rowIndex + rows.length - 1);
        const endColumnIndex = Math.max(
          columnIndex,
          columnIndex + (rows.length && rows[0].length - 1)
        );

        dispatch({
          type: ACTION_TYPE.PASTE,
          id,
          rows,
          activeCell,
          selection
        });

        /* Should select */
        if (rowIndex === endRowIndex && columnIndex === endColumnIndex) return;

        currentGrid.current?.setSelections([
          {
            bounds: {
              top: rowIndex,
              left: columnIndex,
              bottom: endRowIndex,
              right: endColumnIndex
            }
          }
        ]);
      },
      []
    );

    /**
     * Handle cut event
     */
    const handleCut = useCallback((id: SheetID, selection: SelectionArea) => {
      dispatch({
        type: ACTION_TYPE.REMOVE_CELLS,
        id,
        activeCell: null,
        selections: [selection]
      });
    }, []);

    /**
     * Insert new row
     */
    const handleInsertRow = useCallback(
      (id: SheetID, activeCell: CellInterface | null) => {
        if (activeCell === null) return;
        dispatch({
          type: ACTION_TYPE.INSERT_ROW,
          id,
          activeCell
        });
      },
      []
    );

    /**
     * Insert new row
     */
    const handleInsertColumn = useCallback(
      (id: SheetID, activeCell: CellInterface | null) => {
        if (activeCell === null) return;
        dispatch({
          type: ACTION_TYPE.INSERT_COLUMN,
          id,
          activeCell
        });
      },
      []
    );

    /* Handle delete row */
    const handleDeleteRow = useCallback(
      (id: SheetID, activeCell: CellInterface | null) => {
        if (activeCell === null) return;
        dispatch({
          type: ACTION_TYPE.DELETE_ROW,
          id,
          activeCell
        });
      },
      []
    );

    /* Handle delete row */
    const handleDeleteColumn = useCallback(
      (id: SheetID, activeCell: CellInterface | null) => {
        if (activeCell === null) return;
        dispatch({
          type: ACTION_TYPE.DELETE_COLUMN,
          id,
          activeCell
        });
      },
      []
    );

    /**
     * Handle keydown events
     */
    const handleKeyDown = useCallback(
      (
        id: SheetID,
        event: React.KeyboardEvent<HTMLDivElement>,
        activeCell: CellInterface | null
      ) => {
        if (!activeCell) return;
        const isMeta = event.metaKey || event.ctrlKey;
        const isShift = event.shiftKey;
        const keyCode = event.which;
        switch (keyCode) {
          case KeyCodes.KEY_B:
            if (!isMeta) return;
            handleFormattingChange(
              FORMATTING_TYPE.BOLD,
              !getCellConfig(id, activeCell)?.bold
            );
            break;

          case KeyCodes.KEY_I:
            if (!isMeta) return;
            handleFormattingChange(
              FORMATTING_TYPE.ITALIC,
              !getCellConfig(id, activeCell)?.italic
            );
            break;

          case KeyCodes.KEY_U:
            if (!isMeta) return;
            handleFormattingChange(
              FORMATTING_TYPE.UNDERLINE,
              !getCellConfig(id, activeCell)?.underline
            );
            break;

          case KeyCodes.KEY_X:
            if (!isMeta || !isShift) return;
            handleFormattingChange(
              FORMATTING_TYPE.STRIKE,
              !getCellConfig(id, activeCell)?.strike
            );
            break;

          case KeyCodes.BACK_SLASH:
            if (!isMeta) return;
            handleClearFormatting();
            event?.preventDefault();
            break;

          case KeyCodes.KEY_L:
          case KeyCodes.KEY_E:
          case KeyCodes.KEY_R:
            if (!isMeta || !isShift) return;
            const align =
              keyCode === KeyCodes.KEY_L
                ? "left"
                : keyCode === KeyCodes.KEY_E
                ? "center"
                : "right";
            handleFormattingChange(FORMATTING_TYPE.HORIZONTAL_ALIGN, align);
            event?.preventDefault();
            break;
        }

        /* Pass it on to undo hook */
        onUndoKeyDown?.(event);
      },
      [getCellConfig]
    );

    /**
     * Update filters views
     */
    const handleChangeFilter = useCallback(
      (
        id: SheetID,
        filterViewIndex: number,
        columnIndex: number,
        filter?: FilterDefinition
      ) => {
        /* Todo, find rowIndex based on filterViewIndex */
        currentGrid.current?.resetAfterIndices?.({ rowIndex: 0 }, false);

        dispatch({
          type: ACTION_TYPE.CHANGE_FILTER,
          id,
          filterViewIndex,
          columnIndex,
          filter
        });
      },
      []
    );

    /**
     * Callback when scale changes
     */
    const handleScaleChange = useCallback(
      value => {
        /* Update grid dimensions */
        currentGrid.current?.resetAfterIndices?.(
          { rowIndex: 0, columnIndex: 0 },
          false
        );
        /* Set scale */
        setScale(value);
      },
      [selectedSheet]
    );

    const handleShowSheet = useCallback((id: SheetID) => {
      dispatch({
        type: ACTION_TYPE.SHOW_SHEET,
        id
      });
    }, []);

    const handleHideSheet = useCallback((id: SheetID) => {
      dispatch({
        type: ACTION_TYPE.HIDE_SHEET,
        id
      });
    }, []);

    const handleProtectSheet = useCallback((id: SheetID) => {
      dispatch({
        type: ACTION_TYPE.PROTECT_SHEET,
        id
      });
    }, []);

    const handleUnProtectSheet = useCallback((id: SheetID) => {
      dispatch({
        type: ACTION_TYPE.UNPROTECT_SHEET,
        id
      });
    }, []);

    return (
      <>
        <Global
          styles={css`
            .rowsncolumns-spreadsheet {
              font-family: ${fontFamily};
            }
            .rowsncolumns-spreadsheet *,
            .rowsncolumns-spreadsheet *:before,
            .rowsncolumns-spreadsheet *:after {
              box-sizing: border-box;
            }
            .rowsncolumns-grid-container:focus {
              outline: none;
            }
          `}
        />

        <Flex
          flexDirection="column"
          flex={1}
          minWidth={0}
          minHeight={minHeight}
          className="rowsncolumns-spreadsheet"
        >
          {showToolbar ? (
            <Toolbar
              wrap={activeCellConfig?.wrap}
              datatype={activeCellConfig?.datatype}
              plaintext={activeCellConfig?.plaintext}
              format={activeCellConfig?.format}
              fontSize={activeCellConfig?.fontSize}
              fontFamily={activeCellConfig?.fontFamily}
              fill={activeCellConfig?.fill}
              bold={activeCellConfig?.bold}
              italic={activeCellConfig?.italic}
              strike={activeCellConfig?.strike}
              underline={activeCellConfig?.underline}
              color={activeCellConfig?.color}
              percent={activeCellConfig?.percent}
              currency={activeCellConfig?.currency}
              verticalAlign={activeCellConfig?.verticalAlign}
              horizontalAlign={activeCellConfig?.horizontalAlign}
              onFormattingChange={handleFormattingChange}
              onFormattingChangeAuto={handleFormattingChangeAuto}
              onFormattingChangePlain={handleFormattingChangePlain}
              onClearFormatting={handleClearFormatting}
              onMergeCells={handleMergeCells}
              frozenRows={currentSheet.frozenRows}
              frozenColumns={currentSheet.frozenColumns}
              onFrozenRowChange={handleFrozenRowChange}
              onFrozenColumnChange={handleFrozenColumnChange}
              onBorderChange={handleBorderChange}
              onRedo={redo}
              onUndo={undo}
              canRedo={canRedo}
              canUndo={canUndo}
              enableDarkMode={enableDarkMode}
              scale={scale}
              onScaleChange={handleScaleChange}
              fontList={fontList}
            />
          ) : null}
          {showFormulabar ? (
            <Formulabar
              value={formulaInput}
              onChange={handleFormulabarChange}
              onKeyDown={handleFormulabarKeydown}
              onFocus={handleFormulabarFocus}
            />
          ) : null}
          <Workbook
            scale={scale}
            StatusBar={StatusBar}
            showTabStrip={showTabStrip}
            isTabEditable={isTabEditable}
            allowNewSheet={allowNewSheet}
            onResize={handleResize}
            formatter={formatter}
            ref={currentGrid}
            onDelete={handleDelete}
            onFill={handleFill}
            onActiveCellValueChange={handleActiveCellValueChange}
            onActiveCellChange={handleActiveCellChange}
            currentSheet={currentSheet}
            selectedSheet={selectedSheet}
            onChangeSelectedSheet={setSelectedSheet}
            onNewSheet={handleNewSheet}
            theme={theme}
            sheets={sheets}
            onChange={handleChange}
            onSheetChange={handleSheetAttributesChange}
            minColumnWidth={minColumnWidth}
            minRowHeight={minRowHeight}
            CellRenderer={CellRenderer}
            HeaderCellRenderer={HeaderCellRenderer}
            onChangeSheetName={handleChangeSheetName}
            onDeleteSheet={handleDeleteSheet}
            onDuplicateSheet={handleDuplicateSheet}
            onScroll={handleScroll}
            onKeyDown={handleKeyDown}
            onPaste={handlePaste}
            onCut={handleCut}
            onInsertRow={handleInsertRow}
            onInsertColumn={handleInsertColumn}
            onDeleteRow={handleDeleteRow}
            onDeleteColumn={handleDeleteColumn}
            CellEditor={CellEditor}
            onSelectionChange={onSelectionChange}
            selectionMode={selectionMode}
            onChangeFilter={handleChangeFilter}
            showStatusBar={showStatusBar}
            selectionPolicy={selectionPolicy}
            ContextMenu={ContextMenu}
            snap={snap}
            Tooltip={Tooltip}
            onHideSheet={handleHideSheet}
            onShowSheet={handleShowSheet}
            onProtectSheet={handleProtectSheet}
            onUnProtectSheet={handleUnProtectSheet}
          />
        </Flex>
      </>
    );
  })
);

export interface SpreadSheetPropsWithTheme extends SpreadSheetProps {
  theme?: ThemeType;
  initialColorMode?: "light" | "dark";
}
const ThemeWrapper: React.FC<SpreadSheetPropsWithTheme &
  RefAttributeSheetGrid> = forwardRef((props, forwardedRef) => {
  const { theme: defaultTheme = theme, initialColorMode, ...rest } = props;
  return (
    <ThemeProvider theme={defaultTheme}>
      <ColorModeProvider value={initialColorMode}>
        <Spreadsheet {...rest} ref={forwardedRef} />
      </ColorModeProvider>
    </ThemeProvider>
  );
});

export default ThemeWrapper;
