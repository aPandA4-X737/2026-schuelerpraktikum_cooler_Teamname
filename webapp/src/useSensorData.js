import { useCallback, useEffect, useState } from "react";

const API_URL = "http://127.0.0.1:8000/data_wsi/";
const POLL_INTERVALL = 10000; // alle 10s neu laden, da laufend Daten reinkommen

/**
 * /data_wsi/ liefert die komplette Historie verschachtelt:
 * { typ: { sensorname: [{ time, pressure, temperature }, …] } }
 *
 * Tabelle und Suche arbeiten mit einer flachen Liste – deshalb wird hier
 * jede Messung zu einer eigenen Zeile mit Name und Typ aufgelöst.
 */
function flachKlopfen(json) {
  if (!json || typeof json !== "object") return [];

  const zeilen = [];

  for (const [type, sensoren] of Object.entries(json)) {
    for (const [name, messungen] of Object.entries(sensoren ?? {})) {
      for (const messung of messungen ?? []) {
        zeilen.push({
          _id: `${type}-${name}-${messung.time}`,
          type,
          name,
          time: messung.time,
          pressure: messung.pressure,
          temperature: messung.temperature,
        });
      }
    }
  }

  return zeilen;
}

export function useSensorData() {
  const [daten, setDaten] = useState([]);
  const [status, setStatus] = useState("laedt"); // "laedt" | "ok" | "fehler"
  const [fehler, setFehler] = useState(null);
  const [laedtGerade, setLaedtGerade] = useState(false);
  const [letzteAktualisierung, setLetzteAktualisierung] = useState(null);

  const laden = useCallback(async (signal) => {
    setLaedtGerade(true);
    try {
      const response = await fetch(API_URL, { signal });

      // 404 heißt hier nur "noch nichts empfangen" – das ist kein Fehler,
      // die Tabelle zeigt dann ihren Leer-Hinweis.
      if (response.status === 404) {
        setDaten([]);
        setStatus("ok");
        setFehler(null);
        setLetzteAktualisierung(Date.now());
        return;
      }

      if (!response.ok) {
        throw new Error(`Server antwortete mit ${response.status}`);
      }
      const json = await response.json();
      setDaten(flachKlopfen(json));
      setStatus("ok");
      setFehler(null);
      setLetzteAktualisierung(Date.now());
    } catch (error) {
      if (error.name === "AbortError") return;
      setStatus("fehler");
      setFehler(error.message);
    } finally {
      setLaedtGerade(false);
    }
  }, []);

  useEffect(() => {
    const controller = new AbortController();
    let timer;

    // Erst laden, dann den nächsten Abruf planen – so überholen sich
    // langsame Anfragen nicht gegenseitig.
    const planen = (verzoegerung) => {
      timer = setTimeout(async () => {
        await laden(controller.signal);
        if (!controller.signal.aborted) planen(POLL_INTERVALL);
      }, verzoegerung);
    };
    planen(0);

    return () => {
      controller.abort();
      clearTimeout(timer);
    };
  }, [laden]);

  // ohne Argument aufrufen – sonst landet z.B. das Klick-Event als Signal im fetch
  const neuLaden = useCallback(() => laden(), [laden]);

  return { daten, status, fehler, laedtGerade, letzteAktualisierung, neuLaden };
}
