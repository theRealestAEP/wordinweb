import React, { useState } from "react";
import { createRoot } from "react-dom/client";
import { DocxView } from "wordinweb";
import "@fontsource/carlito/400.css";
import "@fontsource/carlito/700.css";

// Import the built browser package, as a consumer does.
const params = new URLSearchParams(location.search);
const source = new URL(`./fixtures/layout/${params.get("fixture")}`, import.meta.url).href;
const editable = params.get("editable") === "true";
await document.fonts.load("16px Carlito");

function App() {
  const [status, setStatus] = useState("Loading");
  return (
    <>
      <p role="status">{status}</p>
      <DocxView
        source={source}
        editable={editable}
        onLoad={({ pageCount }) => setStatus(`Loaded ${pageCount} pages`)}
        onError={(error) => setStatus(`Error: ${error.message}`)}
      />
    </>
  );
}

createRoot(document.getElementById("root")!).render(<App />);
