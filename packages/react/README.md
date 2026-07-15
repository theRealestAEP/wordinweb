# WordInWeb

React components for viewing and editing Word `.docx` files in the browser.

```bash
npm install wordinweb react react-dom
```

```tsx
import { DocxView } from "wordinweb";

export function Preview() {
  return <DocxView source="/report.docx" style={{ height: "100vh" }} />;
}
```

Documentation and source: https://github.com/theRealestAEP/wordinweb

Licensed for noncommercial use under PolyForm Noncommercial 1.0.0.
