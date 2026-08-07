import "./stylesheet.css";
import { useEffect, useState } from "react";

const API_URL = "http://127.0.0.1:8000";
const POLL_INTERVALL = 10000; // alle 10s neu laden, da laufend Daten reinkommen

function formatNumber(value) {
    if (value === null || value === undefined) return "–";
    return Number(value).toFixed(2);
}

/* Die Zeitstempel kommen in zwei Formaten aus dem Backend:
   "20260805_132116" und "2026-08-04T09:27:12.885375". */
function parseTime(time) {
    if (!time) return null;

    const match = /^(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})$/.exec(time);
    if (match) {
        const [, year, month, day, hour, minute, second] = match;
        return new Date(`${year}-${month}-${day}T${hour}:${minute}:${second}`);
    }

    const parsed = new Date(time);
    return isNaN(parsed.getTime()) ? null : parsed;
}

function formatTime(time) {
    const date = parseTime(time);
    if (!date) return time ?? "–";
    return date.toLocaleString("de-DE");
}

/* Sortiert nach Namen, wobei Zahlen als Zahlen verglichen werden.
   Dadurch kommt thruster_2.a vor thruster_10.a und nicht danach. */
function sortByName(sensors) {
    return [...sensors].sort((a, b) =>
        a.name.localeCompare(b.name, "de", { numeric: true })
    );
}

/* Neueste Messung zuerst. Eintraege ohne lesbare Zeit landen am Ende. */
function sortByTime(sensors) {
    return [...sensors].sort((a, b) => {
        const timeA = parseTime(a.time);
        const timeB = parseTime(b.time);
        if (!timeA) return 1;
        if (!timeB) return -1;
        return timeB - timeA;
    });
}

/* /data_wsi/{name} liefert die Historie ebenfalls nach Typ und Name verschachtelt:
   { "thruster": { "thruster_1.a": [ { time, pressure, temperature }, ... ] } }
   Fuer das Fenster brauchen wir nur die Messwerte des gesuchten Sensors. */
function flattenHistory(json, sensorName) {
    return Object.values(json ?? {}).flatMap(
        (sensors) => sensors?.[sensorName] ?? []
    );
}

function SensorDetail({ sensorName, onClose }) {
    const [readings, setReadings] = useState([]);
    const [status, setStatus] = useState("loading");

    // Fenster laesst sich auch mit der Escape-Taste schliessen
    useEffect(() => {
        const handleKeyDown = (event) => {
            if (event.key === "Escape") onClose();
        };
        window.addEventListener("keydown", handleKeyDown);
        return () => window.removeEventListener("keydown", handleKeyDown);
    }, [onClose]);

    useEffect(() => {
        // Wird waehrend des Ladens ein anderer Sensor geklickt, brechen wir die
        // alte Anfrage ab, damit sie das neue Ergebnis nicht ueberschreibt.
        const controller = new AbortController();

        const fetchHistory = async () => {
            setStatus("loading");
            try {
                const response = await fetch(
                    `${API_URL}/data_wsi/${encodeURIComponent(sensorName)}`,
                    { signal: controller.signal }
                );

                if (!response.ok) {
                    setReadings([]);
                    setStatus(response.status === 404 ? "empty" : "error");
                    return;
                }

                const json = await response.json();
                setReadings(flattenHistory(json, sensorName));
                setStatus("ready");
            } catch (error) {
                if (error.name !== "AbortError") setStatus("error");
            }
        };

        fetchHistory();
        return () => controller.abort();
    }, [sensorName]);

    const sortedReadings = sortByTime(readings);

    return (
        <div className="modal-backdrop" onClick={onClose}>
            {/* Klick im Fenster soll es nicht gleich wieder schliessen */}
            <div className="modal" onClick={(event) => event.stopPropagation()}>
                <div className="modal-header">
                    <div>
                        <h2>{sensorName}</h2>
                        <p className="modal-subtitle">
                            {status === "loading"
                                ? "Lade Messungen …"
                                : `${sortedReadings.length} Messungen`}
                        </p>
                    </div>
                    <button className="modal-close" onClick={onClose}>✕</button>
                </div>

                <div className="table-wrapper">
                    <table>
                        <thead>
                            <tr>
                                <th>Zeit</th>
                                <th className="num">Pressure</th>
                                <th className="num">Temperature</th>
                            </tr>
                        </thead>
                        <tbody>
                            {status !== "ready" && (
                                <tr>
                                    <td className="empty" colSpan={3}>
                                        {status === "loading" && "Lade Daten …"}
                                        {status === "empty" && "Keine Daten vorhanden"}
                                        {status === "error" && "Daten konnten nicht geladen werden"}
                                    </td>
                                </tr>
                            )}

                            {status === "ready" && sortedReadings.length === 0 && (
                                <tr>
                                    <td className="empty" colSpan={3}>Keine Daten vorhanden</td>
                                </tr>
                            )}

                            {sortedReadings.map((reading, index) => (
                                <tr key={reading._id ?? index}>
                                    <td>{formatTime(reading.time)}</td>
                                    <td className="num">{formatNumber(reading.pressure)}</td>
                                    <td className="num">{formatNumber(reading.temperature)}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
}

function SensorSection({ title, sensors, onSelectSensor }) {
    return (
        <section className="sensor-section">
            <h2>{title}</h2>
            <div className="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th>Name</th>
                            <th className="num">Pressure</th>
                            <th className="num">Temperature</th>
                        </tr>
                    </thead>
                    <tbody>
                        {sensors.length === 0 && (
                            <tr>
                                <td className="empty" colSpan={3}>Keine Daten vorhanden</td>
                            </tr>
                        )}

                        {sensors.map((sensor, index) => (
                            <tr key={sensor._id ?? index}>
                                <td>
                                    <button
                                        className="sensor-name"
                                        onClick={() => onSelectSensor(sensor.name)}
                                    >
                                        {sensor.name}
                                    </button>
                                </td>
                                <td className="num">{formatNumber(sensor.pressure)}</td>
                                <td className="num">{formatNumber(sensor.temperature)}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
        </section>
    );
}

/* /data/current liefert die Daten verschachtelt nach Typ und Name:
   { "thruster": { "thruster_1.a": { time, pressure, temperature } }, ... }
   Die Tabelle arbeitet mit einer flachen Liste, also bauen wir sie um. */
function flattenCurrent(json) {
    return Object.entries(json ?? {}).flatMap(([type, sensors]) =>
        Object.entries(sensors ?? {}).map(([name, measurement]) => ({
            type,
            name,
            ...measurement,
        }))
    );
}

function SensorTable() {
    const [selectedSensor, setSelectedSensor] = useState(null);

    const [thrusterData, setThrusterData] = useState([]);
    const [oxygenTankData, setOxygenTankData] = useState([]);
    const [hydrogenTankData, setHydrogenTankData] = useState([]);

    useEffect(() => {
        const controller = new AbortController();
        let timer;

        const laden = async () => {
            try {
                const response = await fetch(`${API_URL}/data/current`, {
                    signal: controller.signal,
                });

                // Bei 404 (noch keine Sensordaten) oder einem Serverfehler die
                // bisher angezeigten Werte stehen lassen statt sie zu leeren.
                if (!response.ok) return;

                const json = await response.json();
                const sensors = flattenCurrent(json);

                setThrusterData(sortByName(sensors.filter((item) => item.type === "thruster")));
                setOxygenTankData(sortByName(sensors.filter((item) => item.type === "gas_valve" && item.name.startsWith("o"))));
                setHydrogenTankData(sortByName(sensors.filter((item) => item.type === "gas_valve" && item.name.startsWith("h"))));
            } catch (error) {
                // Backend nicht erreichbar: alte Werte behalten, beim naechsten
                // Durchlauf wird es wieder versucht.
                if (error.name !== "AbortError") console.error(error);
            }
        };

        // Erst laden, dann den naechsten Abruf planen – so ueberholen sich
        // langsame Anfragen nicht gegenseitig.
        const planen = (verzoegerung) => {
            timer = setTimeout(async () => {
                await laden();
                if (!controller.signal.aborted) planen(POLL_INTERVALL);
            }, verzoegerung);
        };
        planen(0);

        return () => {
            controller.abort();
            clearTimeout(timer);
        };
    }, []);

    return (
        <>
            <SensorSection title="Thruster" sensors={thrusterData} onSelectSensor={setSelectedSensor} />
            <SensorSection title="Oxygen Tanks" sensors={oxygenTankData} onSelectSensor={setSelectedSensor} />
            <SensorSection title="Hydrogen Tanks" sensors={hydrogenTankData} onSelectSensor={setSelectedSensor} />

            {selectedSensor && (
                <SensorDetail
                    sensorName={selectedSensor}
                    onClose={() => setSelectedSensor(null)}
                />
            )}
        </>
    );
}

export default SensorTable;
