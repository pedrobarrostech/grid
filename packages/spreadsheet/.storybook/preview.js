import React from "react";
import { addDecorator } from "@storybook/react";
import { withKnobs } from "@storybook/addon-knobs";
import { theme, ThemeProvider } from "@chakra-ui/core";

export const ThemeDecorator = props => {
  return <ThemeProvider theme={theme}>{props.children}</ThemeProvider>;
};

addDecorator(storyFn => <ThemeDecorator>{storyFn()}</ThemeDecorator>);
