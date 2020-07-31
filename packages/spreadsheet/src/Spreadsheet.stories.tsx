import React from "react";
import SpreadSheet, { defaultSheets } from "./Spreadsheet";

export default {
  title: "SpreadSheet",
  component: SpreadSheet
};

export const Default = () => {
  return <SpreadSheet minHeight={800} initialSheets={defaultSheets} />;
};
