import * as React from "react";
import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { MemoryRouter, Routes, Route, Navigate } from "react-router-dom";
import MainPage from "./pages/main";
import EditorPage from "./pages/editor";

/* global document, window, Office, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  const initialPath = window.location.hash.replace("#", "/") || "/main";

  root?.render(
    <FluentProvider theme={webLightTheme} style={{ height: "100%" }}>
      <MemoryRouter initialEntries={[initialPath]}>
        <Routes>
          <Route path="/main" element={<MainPage />} />
          <Route path="/editor" element={<EditorPage />} />
          <Route path="*" element={<Navigate to="/main" replace />} />
        </Routes>
      </MemoryRouter>
    </FluentProvider>
  );
});
