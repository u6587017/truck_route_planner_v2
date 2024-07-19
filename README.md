# Truck Route Planner Version 2
แอปพลิเคชันนี้ช่วยให้ผู้ใช้สามารถเลือกไฟล์ Excel ที่มีข้อมูลคำสั่งซื้อ ประมวลผลข้อมูลเพื่อกำหนดเส้นทางที่มีประสิทธิภาพที่สุด และแสดงผลเส้นทางเหล่านี้บนแผนที่ นอกจากนี้ผู้ใช้ยังสามารถส่งออกเส้นทางที่วางแผนไว้ไปยังไฟล์ Excel ได้ ซึ่งเป็นการใช้ OpenRouteService (ORS) API มาใช้ในการคำนวณระยะทางและ ETA
โดยการใช้งานสามารถเข้าได้ด้วยไฟล์ exe ดังรูปนี้
![icon](https://github.com/user-attachments/assets/151ccd08-ee57-4d4c-8076-f6ade0807c13)<br/>
![prog](https://github.com/user-attachments/assets/56fc63fc-436e-42b0-af8f-40dfb1294b26)
<br />
## การใช้งาน ต้องตามตามลำดับขึ้นดังนี้
### 1 OpenRouteService API Key
ให้ผู้ใช้ใส่ API Key ของ OpenRouteService โดยเข้าไปที่ลิงค์นี้ https://api.openrouteservice.org/ และทำการสมัครโดยสามารถหา API Key ได้ดังภาพด้านล่างนี้
![key](https://github.com/user-attachments/assets/48e41a09-fccb-46f4-a97e-4983b5d5779b)<br/>
### 2 Select Excel:
- โหลดข้อมูลจากไฟล์ Excel .xlxs ซึ่งจะแสดงชื่อไฟล์ที่ทำการเลือกปัจจุบันด้านล่าง จากนั้นจะมีการการแสดงผลบนแผนที่เป็นไฟล์ HTML ซึ่งจะแสดงเส้นทางรถบรรทุกที่วางแผนไว้บนแผนที่มีการแสดงผลเส้นทางด้วยสี โดยแต่ละสีคือรถบรรทุกแต่ละคัน
![key2](https://github.com/user-attachments/assets/74576b16-e713-443d-b99f-0bdb2a479fdb)<br/>
### 3 Export Excel: 
- Export ไฟล์เป็น excel โดยในแต่ละ sheet จะเป็น order ของรถบรรทุกแต่ละคัน
### 4 CtkOptionMenu:
- หลังจากเลือกแล้ว
