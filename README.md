![key](https://github.com/user-attachments/assets/dc45efc5-c3eb-4fcb-ba78-26d83bda6d5f)# Truck Route Planner Version 2
แอปพลิเคชันนี้ช่วยให้ผู้ใช้สามารถเลือกไฟล์ Excel ที่มีข้อมูลคำสั่งซื้อ ประมวลผลข้อมูลเพื่อกำหนดเส้นทางที่มีประสิทธิภาพที่สุด และแสดงผลเส้นทางเหล่านี้บนแผนที่ นอกจากนี้ผู้ใช้ยังสามารถส่งออกเส้นทางที่วางแผนไว้ไปยังไฟล์ Excel ได้ ซึ่งเป็นการใช้ OpenRouteService (ORS) API มาใช้ในการคำนวณระยะทางและ ETA

## การใช้งาน ต้องตามตามลำดับขึ้นดังนี้
### 1 OpenRouteService API Key
ให้ผู้ใช้ใส่ API Key ของ OpenRouteService โดยเข้าไปที่ลิงค์นี้ https://api.openrouteservice.org/
![key](https://github.com/user-attachments/assets/48e41a09-fccb-46f4-a97e-4983b5d5779b)
### 2 Select Excel:
- โหลดข้อมูลจากไฟล์ Excel .xlxs จะมีการการแสดงผลบนแผนที่เป็นไฟล์ HTML: แสดงเส้นทางรถบรรทุกที่วางแผนไว้บนแผนที่มีการแสดงผลเส้นทางด้วยสี โดยแต่ละสีคือจำนวนของรถบรรทุก
### 3 Export Excel: 
- Export ไฟล์เป็น excel โดยในแต่ละ sheet จะเป็น order ของรถบรรทุกแต่ละคัน
