// Anzahl der analogen Eing�nge (einstellbar)
const int numInputs = 5;  // 1 bis n Eing�nge
const int inputPins[] = {A0, A1, A2, A3, A4};  // Entsprechend der Anzahl anpassen

void setup() {
  Serial.begin(9600); // Serielle Verbindung starten
  for (int i = 0; i < numInputs; i++) {
    pinMode(inputPins[i], INPUT); // Pins als Eing�nge definieren
  }
  delay(1000); // Sicherstellen, dass die Verbindung vollst�ndig hergestellt ist
}

void loop() {
  // Pr�fen, ob eine Nachricht empfangen wurde
  if (Serial.available() > 0) {
    handleSerialMessage();
  } else {
    // Kontinuierlich analoge Werte senden
    sendAnalogValues();
  }
  delay(100); // Pause zwischen den Sendungen
}

// Funktion zur Verarbeitung von seriellen Nachrichten
void handleSerialMessage() {
  String command = Serial.readStringUntil('\n'); // Nachricht lesen
  command.trim(); // Leerzeichen entfernen

  if (command == "HELLO_MIXER") {
    Serial.println("MIXER_READY"); // Antwort senden
  }
}

// Funktion zum Senden der analogen Werte
void sendAnalogValues() {
  for (int i = 0; i < numInputs; i++) {
    int sensorValue = analogRead(inputPins[i]); // Wert vom aktuellen Pin lesen
    Serial.print(sensorValue); // Wert senden
    if (i < numInputs - 1) {
      Serial.print(" | "); // Trennzeichen
    }
  }
  Serial.println(); // Zeilenumbruch am Ende der Ausgabe
}
