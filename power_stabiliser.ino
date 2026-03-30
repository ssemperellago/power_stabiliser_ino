// Power Stabilizer for Light Source
// Arduino UNO
// Controls an attenuator to keep photodiode signal constant

// Pin definitions
const int attenuatorPin = A0;     // Analog output to attenuator (A0)
const int photodiodePin = A1;     // Analog input from photodiode (A2)

// Control parameters
int attenuatorValue = 2048;       // Start at half (0-4095 for 12-bit)
int referenceValue = 0;           // Reference photodiode voltage
const int tolerance = 1;         // Acceptable deviation (12-bit scale)
const int step = 1;              // Step size for adjustment (12-bit scale)

#if defined(__AVR__)
  const int DAC_MAX = 255;
  const int CAL_STEP = 4;         // ~64 calibration points over 8-bit range
#else
  const int DAC_MAX = 4095;
  const int CAL_STEP = 64;        // ~64 calibration points over 12-bit range
#endif

const int CAL_SETTLE_MS = 10;     // ms to wait after each attenuator step

void calibrate() {
  Serial.println("# Starting calibration sweep...");
  Serial.println("# attenuator,photodiode");

  int savedAttenuator = attenuatorValue;
  bool halfFound = false;

  analogWrite(attenuatorPin, 0); // Start with no attenuation
  delay(CAL_SETTLE_MS);
  int maxVal = analogRead(photodiodePin);

  for (int val = 0; val <= DAC_MAX; val += CAL_STEP) {
    analogWrite(attenuatorPin, val);
    delay(CAL_SETTLE_MS);
    int pd = analogRead(photodiodePin);
    Serial.print(val);
    Serial.print(",");
    Serial.println(pd);
    // Save attenuator and reference only the first time PD drops below half-max
    if (!halfFound && pd < maxVal / 2) {
      savedAttenuator = val;
      referenceValue = pd;
      halfFound = true;
      Serial.print("# Reference updated to: ");
      Serial.println(referenceValue);
    }
  }

  // Restore previous attenuator value
  analogWrite(attenuatorPin, savedAttenuator);
  attenuatorValue = savedAttenuator;

  Serial.println("# Calibration complete. Attenuator restored.");
}

void setup() {
  analogReadResolution(14); //change to 14-bit resolution
  pinMode(attenuatorPin, OUTPUT);
  pinMode(photodiodePin, INPUT);
  Serial.begin(9600);
#if defined(__AVR__)
  // AVR boards do not support analogWriteResolution
  // Use default 8-bit PWM
#else
  analogWriteResolution(12); // Set 12-bit resolution if supported
#endif
  analogWrite(attenuatorPin, attenuatorValue);
  delay(100); // Wait for system to settle
  referenceValue = analogRead(photodiodePin);
  Serial.print("Reference set to: ");
  Serial.println(referenceValue);
  Serial.println("Send 'set' to update reference to current photodiode value.");
  Serial.println("Send 'cal' to run attenuator calibration sweep.");
}

void loop() {
  int currentValue = analogRead(photodiodePin);
  int error = referenceValue - currentValue;

  
  // Serial commands
  if (Serial.available()) {
    String cmd = Serial.readStringUntil('\n');
    cmd.trim();
    if (cmd.equalsIgnoreCase("set")) {
      referenceValue = currentValue;
      Serial.print("New reference set to: ");
      Serial.println(referenceValue);
    } else if (cmd.equalsIgnoreCase("cal")) {
      calibrate();
    } else if (cmd.equalsIgnoreCase("get")) {
      Serial.print("Photodiode: ");
      Serial.print(currentValue);
      Serial.print(" | Reference: ");
      Serial.print(referenceValue);
      Serial.print(" | Attenuator: ");
      Serial.println(attenuatorValue);
    } else if (cmd.equalsIgnoreCase("setZeroAttenuation")) {
      analogWrite(attenuatorPin, 0);
      delay(10);
      Serial.print("Set to zero attenuation, photodiode value:");
      int pdVal = analogRead(photodiodePin);
      Serial.println(pdVal);
    }
  } // end Serial.available()

  if (abs(error) > tolerance) {
    attenuatorValue += (error > 0) ? -step : step;
    attenuatorValue = constrain(attenuatorValue, 0, DAC_MAX);
    analogWrite(attenuatorPin, attenuatorValue);
  }

  delay(10);
}
