import { useMemo, useState } from "react";
import "./stylesheet.css";

const MINUTE_MS = 60_000;
const MAX_PUNKTE = 30; // mehr Minuten werden im kleinen Diagramm unlesbar

/* Zeitstempel kommen als ISO ("2026-08-06T14:29:21") aus dem Backend,
   ältere Datensätze auch als "20260806_142921". */
function parseZeit(wert) {
  const treffer = /^(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})$/.exec(wert ?? "");
  if (treffer) {
    const [, jahr, monat, tag, stunde, minute, sekunde] = treffer;
    return new Date(jahr, monat - 1, tag, stunde, minute, sekunde);
  }
  const datum = new Date(wert ?? "");
  return Number.isNaN(datum.getTime()) ? null : datum;
}

/**
 * Fasst die Messungen einer Sensorgruppe minutenweise zusammen: pro angefangener
 * Minute ein Punkt, dessen Wert der Durchschnitt aller Sensoren der Gruppe in
 * dieser Minute ist.
 */
function minutenMittel(daten, gehoertDazu) {
  const eimer = new Map();

  for (const zeile of daten) {
    if (!gehoertDazu(zeile)) continue;

    const zeit = parseZeit(zeile.time);
    if (!zeit) continue;

    const beginn = Math.floor(zeit.getTime() / MINUTE_MS) * MINUTE_MS;
    let eintrag = eimer.get(beginn);
    if (!eintrag) {
      eintrag = { beginn, tempSumme: 0, tempAnzahl: 0, druckSumme: 0, druckAnzahl: 0 };
      eimer.set(beginn, eintrag);
    }

    if (typeof zeile.temperature === "number") {
      eintrag.tempSumme += zeile.temperature;
      eintrag.tempAnzahl += 1;
    }
    if (typeof zeile.pressure === "number") {
      eintrag.druckSumme += zeile.pressure;
      eintrag.druckAnzahl += 1;
    }
  }

  // Nur die jüngsten Minuten zeigen – das Fenster wandert mit den Daten mit,
  // damit auch ältere Datensätze sichtbar bleiben.
  return [...eimer.values()]
    .sort((a, b) => a.beginn - b.beginn)
    .slice(-MAX_PUNKTE)
    .map((eintrag) => ({
      zeit: eintrag.beginn,
      anzahl: Math.max(eintrag.tempAnzahl, eintrag.druckAnzahl),
      temperature: eintrag.tempAnzahl ? eintrag.tempSumme / eintrag.tempAnzahl : null,
      pressure: eintrag.druckAnzahl ? eintrag.druckSumme / eintrag.druckAnzahl : null,
    }));
}

/** Achsenbeschriftung auf runde Zahlen bringen (0 / 2,5 / 5 …). */
function achsenRaster(min, max) {
  if (!Number.isFinite(min) || !Number.isFinite(max)) return { ticks: [], min: 0, max: 1 };
  if (min === max) {
    min -= 1;
    max += 1;
  }

  const roh = (max - min) / 3;
  const groesse = 10 ** Math.floor(Math.log10(roh));
  const schritt = ([1, 2, 2.5, 5, 10].find((f) => f * groesse >= roh) ?? 10) * groesse;

  const start = Math.floor(min / schritt) * schritt;
  const ende = Math.ceil(max / schritt) * schritt;

  const ticks = [];
  for (let wert = start; wert <= ende + schritt / 1000; wert += schritt) {
    ticks.push(Number(wert.toFixed(10)));
  }

  return { ticks, min: start, max: ende };
}

function formatZahl(wert, stellen) {
  return wert.toLocaleString("de-DE", {
    minimumFractionDigits: stellen,
    maximumFractionDigits: stellen,
  });
}

function formatUhrzeit(ms) {
  return new Date(ms).toLocaleTimeString("de-DE", {
    hour: "2-digit",
    minute: "2-digit",
  });
}

/* Die Bodenstation sendet nicht durchgehend – die letzten 30 Minutenwerte
   können also über mehrere Tage verteilt sein. Dann gehört das Datum dazu. */
function formatZeitpunkt(ms, mitDatum) {
  if (!mitDatum) return formatUhrzeit(ms);
  return `${new Date(ms).toLocaleDateString("de-DE", {
    day: "2-digit",
    month: "2-digit",
  })} ${formatUhrzeit(ms)}`;
}

const BREITE = 520;
const HOEHE = 210;
const RAND = { oben: 14, rechts: 18, unten: 28, links: 48 };

/**
 * Liniendiagramm mit einer Serie. Bewusst als reines SVG – die App hat
 * keine Diagramm-Bibliothek und braucht für zwei Kurven auch keine.
 */
function LinienDiagramm({ titel, gruppe, untertitel, einheit, punkte, stellen, farbklasse }) {
  const [aktiverIndex, setAktiverIndex] = useState(null);

  const geometrie = useMemo(() => {
    if (punkte.length === 0) return null;

    const werte = punkte.map((p) => p.wert);
    const raster = achsenRaster(Math.min(...werte), Math.max(...werte));

    const innenBreite = BREITE - RAND.links - RAND.rechts;
    const innenHoehe = HOEHE - RAND.oben - RAND.unten;

    // Gleicher Abstand je Minutenwert statt echter Zeitachse: die Bodenstation
    // sendet mit Pausen, sonst würde eine lange Pause alle übrigen Punkte an
    // den rechten Rand quetschen. Bei nur einem Punkt sitzt er mittig.
    const x = (index) =>
      punkte.length === 1
        ? RAND.links + innenBreite / 2
        : RAND.links + (index / (punkte.length - 1)) * innenBreite;

    const y = (wert) =>
      RAND.oben + (1 - (wert - raster.min) / (raster.max - raster.min)) * innenHoehe;

    const koordinaten = punkte.map((p, i) => ({ ...p, x: x(i), y: y(p.wert) }));
    const linie = koordinaten.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
    const flaeche =
      koordinaten.length > 1
        ? `${linie} L ${koordinaten.at(-1).x} ${HOEHE - RAND.unten} L ${koordinaten[0].x} ${HOEHE - RAND.unten} Z`
        : null;

    return { raster, koordinaten, linie, flaeche, y };
  }, [punkte]);

  if (!geometrie) {
    return (
      <figure className={`diagramm ${farbklasse}`}>
        <figcaption className="diagramm-kopf">
          <span className="diagramm-titel">{titel}</span>
          <span className="diagramm-untertitel">{untertitel}</span>
        </figcaption>
        <p className="diagramm-leer">Noch keine Daten für {gruppe} empfangen.</p>
      </figure>
    );
  }

  const { raster, koordinaten, linie, flaeche } = geometrie;
  const letzter = koordinaten.at(-1);
  const aktiver = aktiverIndex == null ? null : koordinaten[aktiverIndex];

  const mitDatum =
    new Date(koordinaten[0].zeit).toDateString() !== new Date(letzter.zeit).toDateString();

  // Punkt unter dem Zeiger suchen. Die Maus-Position muss dafür von Pixeln in
  // das viewBox-Koordinatensystem umgerechnet werden, weil das SVG skaliert.
  const beiBewegung = (event) => {
    const box = event.currentTarget.getBoundingClientRect();
    const zeigerX = ((event.clientX - box.left) / box.width) * BREITE;

    let naechster = 0;
    for (let i = 1; i < koordinaten.length; i += 1) {
      if (Math.abs(koordinaten[i].x - zeigerX) < Math.abs(koordinaten[naechster].x - zeigerX)) {
        naechster = i;
      }
    }
    setAktiverIndex(naechster);
  };

  const beschreibung = `${titel} ${gruppe}: ${koordinaten.length} Minutenwerte von ${formatZeitpunkt(
    koordinaten[0].zeit,
    mitDatum,
  )} bis ${formatZeitpunkt(letzter.zeit, mitDatum)}, zuletzt ${formatZahl(
    letzter.wert,
    stellen,
  )} ${einheit}.`;

  return (
    <figure className={`diagramm ${farbklasse}`}>
      <figcaption className="diagramm-kopf">
        <div>
          <span className="diagramm-titel">{titel}</span>
          <span className="diagramm-untertitel">{untertitel}</span>
        </div>
        <div className="diagramm-aktuell">
          <span className="diagramm-aktuell-wert">
            {formatZahl(letzter.wert, stellen)}
            <span className="diagramm-einheit">{einheit}</span>
          </span>
          <span className="diagramm-aktuell-label">letzte Minute</span>
        </div>
      </figcaption>

      <div className="diagramm-flaeche">
        <svg
          className="diagramm-svg"
          viewBox={`0 0 ${BREITE} ${HOEHE}`}
          role="img"
          aria-label={beschreibung}
          onPointerMove={beiBewegung}
          onPointerLeave={() => setAktiverIndex(null)}
        >
          {raster.ticks.map((tick) => {
            const y = geometrie.y(tick);
            return (
              <g key={tick}>
                <line
                  className="diagramm-raster"
                  x1={RAND.links}
                  y1={y}
                  x2={BREITE - RAND.rechts}
                  y2={y}
                />
                <text className="diagramm-tick" x={RAND.links - 8} y={y + 4} textAnchor="end">
                  {formatZahl(tick, tick % 1 === 0 ? 0 : stellen)}
                </text>
              </g>
            );
          })}

          {flaeche && <path className="diagramm-wash" d={flaeche} />}
          <path className="diagramm-linie" d={linie} fill="none" />

          {/* Endpunkt bleibt immer sichtbar, der Ring hebt ihn von der Linie ab */}
          <circle className="diagramm-punkt-ring" cx={letzter.x} cy={letzter.y} r={6} />
          <circle className="diagramm-punkt" cx={letzter.x} cy={letzter.y} r={4} />

          {aktiver && (
            <>
              <line
                className="diagramm-fadenkreuz"
                x1={aktiver.x}
                y1={RAND.oben}
                x2={aktiver.x}
                y2={HOEHE - RAND.unten}
              />
              <circle className="diagramm-punkt-ring" cx={aktiver.x} cy={aktiver.y} r={6} />
              <circle className="diagramm-punkt" cx={aktiver.x} cy={aktiver.y} r={4} />
            </>
          )}

          <text
            className="diagramm-tick"
            x={koordinaten[0].x}
            y={HOEHE - 8}
            textAnchor="start"
          >
            {formatZeitpunkt(koordinaten[0].zeit, mitDatum)}
          </text>
          {koordinaten.length > 1 && (
            <text className="diagramm-tick" x={letzter.x} y={HOEHE - 8} textAnchor="end">
              {formatZeitpunkt(letzter.zeit, mitDatum)}
            </text>
          )}
        </svg>

        {aktiver && (
          <div
            className="diagramm-tooltip"
            style={{
              left: `${(aktiver.x / BREITE) * 100}%`,
              transform: `translate(${aktiver.x > BREITE / 2 ? "-100%" : "0"}, 0)`,
            }}
          >
            <span className="diagramm-tooltip-zeit">
              {formatZeitpunkt(aktiver.zeit, mitDatum)}
            </span>
            <span className="diagramm-tooltip-wert">
              {formatZahl(aktiver.wert, stellen)} {einheit}
            </span>
            <span className="diagramm-tooltip-info">
              Ø aus {aktiver.anzahl} {aktiver.anzahl === 1 ? "Messung" : "Messungen"}
            </span>
          </div>
        )}
      </div>
    </figure>
  );
}

/**
 * Zwei Verlaufsdiagramme einer Sensorgruppe: Temperatur und Druck, jeweils als
 * Minutendurchschnitt über alle Sensoren der Gruppe.
 */
export default function Trenddiagramme({ daten, gruppe, gehoertDazu }) {
  const minuten = useMemo(() => minutenMittel(daten, gehoertDazu), [daten, gehoertDazu]);

  const temperaturPunkte = useMemo(
    () =>
      minuten
        .filter((m) => m.temperature != null)
        .map((m) => ({ zeit: m.zeit, wert: m.temperature, anzahl: m.anzahl })),
    [minuten],
  );

  const druckPunkte = useMemo(
    () =>
      minuten
        .filter((m) => m.pressure != null)
        .map((m) => ({ zeit: m.zeit, wert: m.pressure, anzahl: m.anzahl })),
    [minuten],
  );

  return (
    <section className="diagramme" aria-label={`Verlauf der Messwerte: ${gruppe}`}>
      <LinienDiagramm
        titel="Ø Temperatur"
        gruppe={gruppe}
        untertitel={`Ø je Minute in K · letzte ${MAX_PUNKTE} Minutenwerte`}
        einheit="K"
        punkte={temperaturPunkte}
        stellen={1}
        farbklasse="diagramm--temperatur"
      />
      <LinienDiagramm
        titel="Ø Druck"
        gruppe={gruppe}
        untertitel={`Ø je Minute in bar · letzte ${MAX_PUNKTE} Minutenwerte`}
        einheit="bar"
        punkte={druckPunkte}
        stellen={2}
        farbklasse="diagramm--druck"
      />
    </section>
  );
}
