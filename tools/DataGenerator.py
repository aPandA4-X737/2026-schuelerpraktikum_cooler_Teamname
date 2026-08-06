import datetime
import json
import math
import os.path
import random
import time
from collections import Counter

BASE_PATH = path = os.path.dirname(os.path.dirname(__file__))

# Erlaubte Wertebereiche. Muessen mit den Filtergrenzen in Bodenstation.py uebereinstimmen.
MIN_PRESSURE, MAX_PRESSURE = 0.5, 9.0
MIN_TEMPERATURE, MAX_TEMPERATURE = 200.0, 500.0

# Art des erzeugten Datensatzes
RESULT_NORMAL = 1             # sauberer Wert auf der Sinuskurve
RESULT_OUT_OF_RANGE = 2       # Werte weit ausserhalb der erlaubten Grenzen
RESULT_WRONG_TYPES = 3        # falsche Datentypen
RESULT_WRONG_FILE_FORMAT = 4  # .txt / .yaml / .xml / .pdf statt .json
RESULT_OFF_PATTERN = 5        # innerhalb der Grenzen, aber deutlich neben der Kurve

RESULT_NAMES = {
    RESULT_NORMAL: "normal",
    RESULT_OUT_OF_RANGE: "out_of_range",
    RESULT_WRONG_TYPES: "wrong_types",
    RESULT_WRONG_FILE_FORMAT: "wrong_file_format",
    RESULT_OFF_PATTERN: "off_pattern",
}


def sine_value(t, period, phase, minimum, maximum, margin=0.15, noise=0.02):
    """Wert auf einer Sinuskurve zwischen minimum und maximum, mit leichtem Rauschen.

    margin haelt die Kurve von den Grenzen weg, damit Rauschen und Ausreisser
    darunter noch Platz haben.
    """
    center = (minimum + maximum) / 2
    amplitude = (maximum - minimum) / 2 * (1 - margin)
    value = center + amplitude * math.sin(2 * math.pi * t / period + phase)
    value += random.uniform(-1, 1) * noise * (maximum - minimum)
    return min(max(value, minimum), maximum)


def off_pattern_value(expected, minimum, maximum):
    """Wert innerhalb der Grenzen, aber deutlich neben dem erwarteten Kurvenwert."""
    offset = random.uniform(0.30, 0.50) * (maximum - minimum)

    # Nur die Richtungen zulassen, die noch innerhalb der Grenzen liegen.
    candidates = []
    if expected + offset <= maximum:
        candidates.append(expected + offset)
    if expected - offset >= minimum:
        candidates.append(expected - offset)

    if not candidates:  # sollte nicht vorkommen, da offset <= (maximum - minimum) / 2
        return min(max(expected + offset, minimum), maximum)

    return random.choice(candidates)


class SensorKey:
    """Unique key of a sensor"""

    def __init__(self, name: str, type: str):
        """Constructor"""
        self.name: str = name
        self.type: str = type

        # Feste Kurve pro Sensor: aus dem Namen abgeleitet, damit sie einen
        # Neustart des Generators ueberlebt.
        curve_rng = random.Random(name)
        self.phase: float = curve_rng.uniform(0, 2 * math.pi)
        self.pressure_period: float = curve_rng.uniform(40.0, 90.0)      # Sekunden pro Schwingung
        self.temperature_period: float = curve_rng.uniform(60.0, 120.0)


class Sensor:
    """Sensor object, which stores all information of a given sensor."""

    def __init__(self, name: str, type: str, pressure: float | None, temperature: float | None):
        """Constructor"""

        self.name: str = name
        self.type: str = type
        self.pressure: float | None = pressure
        self.temperature: float | None = temperature


class DataGenerator:
    """Data Generator, which provides and stores sensor data of a given satellite."""

    def __init__(self):
        """Constructor"""
        self.number=1
        self.start_time = time.monotonic()
        self.available_sensors: list[SensorKey] = [
            SensorKey(name="thruster_1.a", type="thruster"),
            SensorKey(name="thruster_1.b", type="thruster"),
            SensorKey(name="thruster_1.c", type="thruster"),
            SensorKey(name="thruster_2.a", type="thruster"),
            SensorKey(name="thruster_2.b", type="thruster"),
            SensorKey(name="thruster_2.c", type="thruster"),
            SensorKey(name="thruster_3.a", type="thruster"),
            SensorKey(name="thruster_3.b", type="thruster"),
            SensorKey(name="thruster_3.c", type="thruster"),
            SensorKey(name="oxygen_tank_1", type="gas_valve"),
            SensorKey(name="oxygen_tank_2", type="gas_valve"),
            SensorKey(name="hydrogen_tank_1", type="gas_valve"),
            SensorKey(name="hydrogen_tank_2", type="gas_valve")
        ]

    def generate_new_sensor_data(self, result):
        selected_key_idx = random.randint(0, len(self.available_sensors) - 1)
        selected_key = self.available_sensors[selected_key_idx]

        # Zeit seit dem Start des Generators. Bewusst die Uhr und kein Zaehler,
        # damit die Kurve auch bei geaenderter Schlafzeit stimmt.
        t = time.monotonic() - self.start_time

        pressure = sine_value(
            t, selected_key.pressure_period, selected_key.phase,
            MIN_PRESSURE, MAX_PRESSURE
        )
        temperature = sine_value(
            t, selected_key.temperature_period, selected_key.phase,
            MIN_TEMPERATURE, MAX_TEMPERATURE
        )

        if result == RESULT_OUT_OF_RANGE:
            pressure = random.uniform(1473829.123, 3847530.283)
            temperature = random.uniform(-220, -100)
        elif result == RESULT_OFF_PATTERN:
            # Mal nur der Druck, mal nur die Temperatur, mal beide.
            broken = random.choice(["pressure", "temperature", "both"])
            if broken in ("pressure", "both"):
                pressure = off_pattern_value(pressure, MIN_PRESSURE, MAX_PRESSURE)
            if broken in ("temperature", "both"):
                temperature = off_pattern_value(temperature, MIN_TEMPERATURE, MAX_TEMPERATURE)

        if result == RESULT_WRONG_TYPES:
            match random.randint(0, 2):
                case 0:
                    sensor_data = Sensor(
                        name="Error",
                        type="Error",
                        pressure="Error",
                        temperature="Error"
                    )
                case 1:
                    sensor_data = Sensor(
                        name=selected_key.name,
                        type=selected_key.type,
                        pressure=str(round(pressure, 3)),
                        temperature=temperature
                    )
                case _:
                    sensor_data = Sensor(
                        name=selected_key.name,
                        type=selected_key.type,
                        pressure=pressure,
                        temperature=None
                    )
        else:
            sensor_data = Sensor(
                name=selected_key.name,
                type=selected_key.type,
                pressure=pressure,
                temperature=temperature

            )

        return sensor_data

    def store_sensor_data(self,result,data: Sensor):
        content = data.__dict__
        if result == RESULT_WRONG_FILE_FORMAT:
            match self.number%4:
                case 0:
                    file_name="/data/TM_"+datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".txt"
                case 1:
                    file_name="/data/TM_"+datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".yaml"
                case 2:
                    file_name="/data/TM_"+datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".xml"
                case 3:
                    file_name="/data/TM_"+datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".pdf"
        else:
            file_name = "/data/TM_"+datetime.datetime.now().strftime("%Y%m%d_%H%M%S")+".json"
        os.makedirs(BASE_PATH + "/data", exist_ok=True)
        with open(BASE_PATH + file_name, "w") as file:
            json.dump(content, file)
        self.number+=1

if __name__ == '__main__':

    generator = DataGenerator()
    list1=[]

    try:
        while True:
            result = random.choices(
                [RESULT_NORMAL, RESULT_OUT_OF_RANGE, RESULT_WRONG_TYPES,
                 RESULT_WRONG_FILE_FORMAT, RESULT_OFF_PATTERN],
                weights=[0.70, 0.05, 0.05, 0.05, 0.15]
            )[0]
            data = generator.generate_new_sensor_data(result)
            generator.store_sensor_data(result,data=data)

            print(f"Sucessfully stored: [{RESULT_NAMES[result]:<17}] "
                  f"{data.name:<16} p={data.pressure} T={data.temperature}")
            list1.append(result)
            # Die Dateinamen haben nur Sekundenaufloesung: kuerzer als 1 s
            # wuerde Datensaetze mit gleichem Namen ueberschreiben.
            time.sleep(1.0)
    except KeyboardInterrupt:
        count = Counter(RESULT_NAMES[r] for r in list1)
        print()
        print(f"Insgesamt {len(list1)} Datensaetze erzeugt:")
        for name, amount in count.most_common():
            print(f"  {name:<17} {amount:>4}  ({amount / len(list1):.0%})")
