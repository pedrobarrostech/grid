import React, {
  useRef,
  useCallback,
  useEffect,
  useState,
  useMemo,
} from "react";
import { EditorProps } from "@rowsncolumns/grid/dist/hooks/useEditable";
import { AutoSizerCanvas } from "@rowsncolumns/grid";
import TextEditor from "./Text";
import ListEditor from "./List";
import { useColorMode } from "@chakra-ui/core";
import {
  DARK_MODE_COLOR_LIGHT,
  DEFAULT_FONT_FAMILY,
  cellToAddress,
} from "../constants";
import { EditorType } from "../types";

export interface CustomEditorProps extends EditorProps {
  background?: string;
  color?: string;
  fontSize?: number;
  fontFamily?: string;
  wrap?: any;
  horizontalAlign?: any;
  scale?: number;
  editorType?: EditorType;
  options?: string[];
  underline?: boolean;
}

/**
 * Default cell editor
 * @param props
 */
const Editor: React.FC<CustomEditorProps> = (props) => {
  const {
    rowIndex,
    columnIndex,
    onChange,
    onSubmit,
    onCancel,
    position,
    cell,
    nextFocusableCell,
    value = "",
    activeCell,
    autoFocus = true,
    background,
    color,
    fontSize = 12,
    fontFamily = DEFAULT_FONT_FAMILY,
    wrap: cellWrap = "nowrap",
    selections,
    scrollPosition,
    horizontalAlign,
    underline,
    scale = 1,
    editorType = "text",
    options,
    ...rest
  } = props;
  const wrapping: any = cellWrap === "wrap" ? "wrap" : "nowrap";
  const { colorMode } = useColorMode();
  const isLight = colorMode === "light";
  const backgroundColor =
    background !== void 0
      ? background
      : isLight
      ? "white"
      : DARK_MODE_COLOR_LIGHT;
  const textColor =
    color !== void 0 ? color : isLight ? DARK_MODE_COLOR_LIGHT : "white";
  const borderWidth = 2;
  const padding = 10; /* 2 + 1 + 1 + 2 + 2 */
  const hasScrollPositionChanged = useRef(false);
  const isMounted = useRef(false);
  const textSizer = useRef(AutoSizerCanvas());
  const { x = 0, y = 0, width = 0, height = 0 } = position;
  const getWidth = useCallback(
    (text) => {
      /*  Set font */
      textSizer.current.setFont({
        fontSize,
        fontFamily,
      });

      const textWidth = textSizer.current.measureText(text)?.width || 0;
      return Math.max(textWidth + padding, width + borderWidth / 2);
    },
    [width, fontSize, fontFamily, wrapping]
  );
  /* Keep updating value of input */
  useEffect(() => {
    setInputWidth(getWidth(value));
  }, [value]);
  /* Width of the input  */
  const [inputWidth, setInputWidth] = useState(() => getWidth(value));
  /* Tracks scroll position: To show address token */

  useEffect(() => {
    if (!isMounted.current) return;
    hasScrollPositionChanged.current = true;
  }, [scrollPosition]);
  /* Set mounted state */
  useEffect(() => {
    /* Set mounted ref */
    isMounted.current = true;
  }, []);
  const address = useMemo(
    () => hasScrollPositionChanged.current && cellToAddress(activeCell),
    [activeCell, hasScrollPositionChanged.current]
  );
  const inputHeight = height;
  /* Change */
  const handleChange = useCallback(
    (value) => {
      onChange?.(value, cell);
    },
    [cell]
  );
  /* Submit */
  const handleSubmit = useCallback(
    (value, direction) => {
      const nextCell = direction ? nextFocusableCell(cell, direction) : cell;
      onSubmit?.(value, cell, nextCell);
    },
    [cell]
  );
  /* Cancel */
  const handleCancel = useCallback(() => {
    onCancel?.();
  }, [cell]);
  return (
    <div
      style={{
        top: y,
        left: x,
        position: "absolute",
        width: inputWidth,
        height: inputHeight + borderWidth / 2,
        padding: borderWidth,
        boxShadow: "0 2px 6px 2px rgba(60,64,67,.15)",
        border: "2px #1a73e8 solid",
        background: backgroundColor,
      }}
    >
      {address ? (
        <div
          style={{
            position: "absolute",
            left: -2,
            marginBottom: 4,
            fontSize: 12,
            lineHeight: "14px",
            padding: 6,
            paddingTop: 4,
            paddingBottom: 4,
            boxShadow: "0px 1px 2px rgba(0,0,0,0.5)",
            bottom: "100%",
            background: "#4589eb",
            color: "white",
          }}
        >
          {address}
        </div>
      ) : null}
      {editorType === "text" ? (
        <TextEditor
          value={value}
          fontFamily={fontFamily}
          fontSize={fontSize}
          scale={scale}
          color={textColor}
          wrapping={wrapping}
          horizontalAlign={horizontalAlign}
          underline={underline}
          onChange={handleChange}
          onSubmit={handleSubmit}
          onCancel={handleCancel}
        />
      ) : null}
      {editorType === "list" ? (
        <ListEditor
          value={value}
          fontFamily={fontFamily}
          fontSize={fontSize}
          scale={scale}
          color={textColor}
          wrapping={wrapping}
          horizontalAlign={horizontalAlign}
          onChange={handleChange}
          onSubmit={handleSubmit}
          onCancel={handleCancel}
          options={options}
        />
      ) : null}
    </div>
  );
};

export default Editor;
