String data;

void setup() {
  Serial.begin(115200);
}

void loop() {
  //data = Serial.read();
  //Serial.println(data);
  //int sensorValue = analogRead(A0);
  Serial.println("{1:50,2:100,3:48}");
  delay(1);
}
