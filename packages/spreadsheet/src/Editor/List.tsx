import React, {
  forwardRef,
  useEffect,
  useRef,
  memo,
  useState,
  useCallback
} from "react";
import { KeyCodes, Direction } from "@rowsncolumns/grid";
import Downshift from "downshift";
import { Box, useTheme, useColorMode } from "@chakra-ui/core";
import { DARK_MODE_COLOR } from "../constants";

export interface ListEditorProps {
  options?: string[];
  value: string;
  onChange: (value: string) => void;
  onSubmit: (value: string | undefined, direction?: Direction) => void;
  onCancel: () => void;
  fontFamily: string;
  fontSize: number;
  scale: number;
  color: string;
  wrapping: any;
  horizontalAlign: any;
}
const EMPTY_ARRAY: string[] = [];

export type RefAttribute = {
  ref?: React.MutableRefObject<HTMLInputElement>;
};

const ListEditor: React.FC<ListEditorProps & RefAttribute> = memo(
  forwardRef((props, forwardedRef) => {
    const inputRef = useRef<HTMLInputElement | null>(null);
    const {
      value: initialValue,
      onChange,
      onSubmit,
      onCancel,
      fontFamily,
      fontSize,
      scale,
      color,
      wrapping,
      horizontalAlign,
      options = EMPTY_ARRAY,
      ...rest
    } = props;
    const initialInputValue = useRef(initialValue);
    const direction = useRef<Direction | null>(null);
    useEffect(() => {
      if (!inputRef.current) return;
      inputRef.current.focus();
      /* Focus cursor at the end */
      inputRef.current.selectionStart = initialValue.length;

      if (forwardedRef)
        (forwardedRef as React.MutableRefObject<HTMLInputElement>).current =
          inputRef.current;
    }, []);
    const theme = useTheme();
    const { colorMode } = useColorMode();
    const isLight = colorMode === "light";
    const borderColor = isLight
      ? theme.colors.gray[300]
      : theme.colors.gray[600];
    const inputColor = isLight ? DARK_MODE_COLOR : theme.colors.white;
    const dropdownBgColor = isLight
      ? theme.colors.white
      : theme.colors.gray[700];
    return (
      <>
        <Downshift
          id="list-editor"
          initialInputValue={initialValue}
          onChange={value => {
            onSubmit(value ?? undefined);
          }}
          initialIsOpen
          initialSelectedItem={initialValue}
        >
          {({
            getRootProps,
            getInputProps,
            getMenuProps,
            getItemProps,
            inputValue,
            selectedItem,
            highlightedIndex
          }) => {
            direction.current = null;
            const inputProps = getInputProps();
            const items = options.filter(item => {
              if (!inputValue || inputValue === initialInputValue.current)
                return true;
              return new RegExp(inputValue, "gi").test(item);
            });
            return (
              <Box {...getRootProps(undefined, { suppressRefError: true })}>
                <input
                  {...inputProps}
                  ref={inputRef}
                  style={{
                    fontFamily,
                    fontSize: fontSize * scale,
                    width: "100%",
                    height: "100%",
                    padding: "0 1px",
                    margin: 0,
                    boxSizing: "border-box",
                    borderWidth: 0,
                    outline: "none",
                    resize: "none",
                    overflow: "hidden",
                    verticalAlign: "top",
                    background: "transparent",
                    color: color,
                    whiteSpace: "pre-wrap",
                    textAlign: horizontalAlign,
                    lineHeight: "14px"
                  }}
                  onChange={(e: React.ChangeEvent<HTMLInputElement>) => {
                    inputProps.onChange?.(e);
                    onChange?.(e.target.value);
                  }}
                  onKeyDown={(e: React.KeyboardEvent<HTMLInputElement>) => {
                    if (!inputRef.current) return;
                    const isShiftKey = e.nativeEvent.shiftKey;
                    const isMetaKey = e.nativeEvent.metaKey;
                    const value = inputRef.current.value;

                    if (
                      inputProps &&
                      inputProps.onKeyDown &&
                      e.which !== KeyCodes.Enter
                    ) {
                      inputProps.onKeyDown(e);
                    }

                    // Enter key
                    if (e.which === KeyCodes.Enter) {
                      /* Add a new line when Cmd/Ctrl key is pressed */
                      // if (isMetaKey) {
                      //   return onChange(value + "\n");
                      // }
                      onSubmit &&
                        onSubmit(
                          highlightedIndex === null
                            ? value
                            : items[highlightedIndex],
                          isShiftKey ? Direction.Up : Direction.Down
                        );
                    }

                    if (e.which === KeyCodes.Escape) {
                      onCancel && onCancel();
                    }

                    if (e.which === KeyCodes.Tab) {
                      onSubmit &&
                        onSubmit(
                          highlightedIndex === null
                            ? value
                            : items[highlightedIndex],
                          isShiftKey ? Direction.Left : Direction.Right
                        );
                    }
                  }}
                />
                <Box
                  {...getMenuProps()}
                  width="auto"
                  left="-2px"
                  shadow="md"
                  background={dropdownBgColor}
                  pb={1}
                  pt={1}
                  position="absolute"
                  top="100%"
                  mt="2px"
                  borderColor={borderColor}
                  borderStyle="solid"
                  borderWidth={1}
                  minWidth="calc(100% + 4px)"
                >
                  {items.map((item, index) => {
                    return (
                      <Box
                        fontSize={12}
                        padding={1}
                        pl={2}
                        pr={2}
                        whiteSpace="nowrap"
                        {...getItemProps({ key: item, item, index })}
                        color={inputColor}
                        cursor="pointer"
                        style={{
                          fontWeight: selectedItem === item ? "bold" : "normal",
                          backgroundColor:
                            highlightedIndex === index || selectedItem === item
                              ? isLight
                                ? theme.colors.gray[100]
                                : "rgba(255,255,255,0.06)"
                              : dropdownBgColor
                        }}
                      >
                        {item}
                      </Box>
                    );
                  })}
                </Box>
              </Box>
            );
          }}
        </Downshift>
      </>
    );
  })
);

export default ListEditor;
