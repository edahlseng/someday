import { CalendarPicker } from "@/components/calendar-picker";
import "./App.css";
import "./index.css";

import { ThemeProvider } from "@/components/theme-provider";
function App() {
  return (
    <ThemeProvider>
      <CalendarPicker />
    </ThemeProvider>
  );
}

export default App;
