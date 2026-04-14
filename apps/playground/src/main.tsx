import * as React from "react";
import { Analytics } from "@vercel/analytics/react";
import { createRoot } from "react-dom/client";
import { App } from "./App";
import { PlaygroundCustomizerProvider } from "./components/playground-customizer";
import { ThemeProvider } from "./components/theme-provider";
import "./index.css";

createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <ThemeProvider attribute="class" defaultTheme="system" enableSystem disableTransitionOnChange>
      <PlaygroundCustomizerProvider>
        <App />
      </PlaygroundCustomizerProvider>
      <Analytics />
    </ThemeProvider>
  </React.StrictMode>
);
