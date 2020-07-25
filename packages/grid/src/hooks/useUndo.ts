import React, { useReducer, useRef, useCallback, useEffect } from "react";
import { KeyCodes } from "../types";

/**
 * Spec: https://tools.ietf.org/html/rfc6902
 *
 * add, { "op": "add", "path": ["data", "1,2"], "value": "hello world" }
 * remove { "op": "remove", "path": ["data", "1,2"], "value": "hello world" }
 * replace { "op": "replace", "path": ["data", "1,2"], "value": "hello world" }
 * move { "op": "move", "from": "/a/b/c", "path": "/a/b/d" }
 * copy
 */

export interface UndoProps {
  id?: number;
  onRedo?: (patches: any) => void;
  onUndo?: (patches: any) => void;
}

export interface UndoManager {
  undo: () => void;
  redo: () => void;
  add: (patches: any) => void;
  canUndo: boolean;
  canRedo: boolean;
  onKeyDown: (e: React.KeyboardEvent<HTMLDivElement>) => void;
}

export interface PatchInterface<T> {
  patches: T;
  inversePatches: T;
}

/**
 * Undo/Redo hook
 * @param
 */
const useUndo = <T>(props: UndoProps = {}): UndoManager => {
  const { onRedo, onUndo } = props;
  const undoStack = useRef<PatchInterface<T>[]>([]);
  const undoStackPointer = useRef<number>(-1);
  const [_, forceRender] = useReducer((s) => s + 1, 0);

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLDivElement>) => {
      const isMeta = e.metaKey || e.ctrlKey;
      const isUndo = isMeta && e.which === KeyCodes.Z;
      const isRedo = (e.shiftKey && isUndo) || e.which === KeyCodes.KEY_Y;
      if (!isRedo && !isUndo) return;
      if (isRedo) {
        handleRedo();
      } else {
        handleUndo();
      }
    },
    []
  );

  const handleUndo = useCallback(() => {
    if (undoStackPointer.current < 0) return;
    const patches = undoStack.current[undoStackPointer.current].inversePatches;
    undoStackPointer.current--;
    onUndo && onUndo(patches);
    forceRender();
  }, []);

  const handleRedo = useCallback(() => {
    if (undoStackPointer.current === undoStack.current.length - 1) return;
    undoStackPointer.current++;
    const patches = undoStack.current[undoStackPointer.current].patches;
    onRedo && onRedo(patches);
    forceRender();
  }, []);

  const addUndoable = useCallback((history: PatchInterface<T>) => {
    const { patches, inversePatches } = history;
    const pointer = ++undoStackPointer.current;
    undoStack.current.length = pointer;
    undoStack.current[pointer] = { patches, inversePatches };
    forceRender();
  }, []);

  return {
    undo: handleUndo,
    redo: handleRedo,
    add: addUndoable,
    onKeyDown: handleKeyDown,
    canUndo: !(undoStackPointer.current < 0),
    canRedo: !(undoStackPointer.current === undoStack.current.length - 1),
  };
};

export default useUndo;
