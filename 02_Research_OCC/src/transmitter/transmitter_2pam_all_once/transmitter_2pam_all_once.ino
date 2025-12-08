const int ledPin = 11; 
const int LEVEL_OFF = 0;
const int LEVEL_ON = 255;
// Delay time in milliseconds
const unsigned long symbolDelay = 66; 

String dataBuffer = "";
bool dataReady = false;
bool sending = false;
int dataIndex = 0;
unsigned long previousMillis = 0;

void setup() {
  Serial.begin(9600);
  pinMode(ledPin, OUTPUT);
  analogWrite(ledPin, LEVEL_OFF);
  dataBuffer.reserve(2048); 
}

void loop() {
  // 1. Receive data from PC
  if (!dataReady) {
    while (Serial.available() > 0) {
      char c = Serial.read();
      
      if (c == '\n') { 
        dataReady = true;
        sending = true;
        dataIndex = 0;
        previousMillis = millis();
        break;
      } else {
        dataBuffer += c; 
      }
    }
  }

  // 2. Blink LED using timer
  if (dataReady && sending) {
    unsigned long currentMillis = millis();

    if (currentMillis - previousMillis >= symbolDelay) {
      previousMillis = currentMillis;

      if (dataIndex < dataBuffer.length()) {
        char symbol = dataBuffer.charAt(dataIndex);

        if (symbol == '1') {
          analogWrite(ledPin, LEVEL_ON);
        } else { 
          analogWrite(ledPin, LEVEL_OFF);
        }
        dataIndex++; 
        
      } else {
        // Finish
        sending = false;
        dataReady = false;
        dataBuffer = "";
        analogWrite(ledPin, LEVEL_OFF);
      }
    }
  }
}