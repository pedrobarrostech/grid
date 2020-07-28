import React, { useEffect, useState } from "react";
import { canUseDOM } from "@rowsncolumns/grid";
const WebFont = canUseDOM ? require("webfontloader") : null;

export interface UseFontResults {
  isFontActive: boolean;
}

export interface UseFontProps {
  google?: string[];
}

const useFonts = (config?: WebFont.Config) => {
  const [isFontActive, setIsFontActive] = useState(false);
  useEffect(() => {
    if (config && WebFont !== null) {
      WebFont.load({
        ...config,
        loading: () => setIsFontActive(false),
        active: () => setIsFontActive(true),
      });
    }
  }, [config]);

  return {
    isFontActive,
  };
};

export default useFonts;
