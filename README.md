# Truck Route Planner Version 2
แอปพลิเคชันนี้ช่วยให้ผู้ใช้สามารถเลือกไฟล์ Excel ที่มีข้อมูลคำสั่งซื้อ ประมวลผลข้อมูลเพื่อกำหนดเส้นทางที่มีประสิทธิภาพที่สุด และแสดงผลเส้นทางเหล่านี้บนแผนที่ นอกจากนี้ผู้ใช้ยังสามารถส่งออกเส้นทางที่วางแผนไว้ไปยังไฟล์ Excel ได้ ซึ่งเป็นการใช้ OpenRouteService (ORS) API มาใช้ในการคำนวณระยะทางและ ETA
โดยการใช้งานสามารถเข้าได้ด้วยไฟล์ exe ดังรูปนี้
![icon](https://github.com/user-attachments/assets/151ccd08-ee57-4d4c-8076-f6ade0807c13)<br/>
![prog](https://github.com/user-attachments/assets/56fc63fc-436e-42b0-af8f-40dfb1294b26)
<br />
## การใช้งาน ต้องตามตามลำดับขึ้นดังนี้
### 1 OpenRouteService API Key
- การจะใช้งานโปรแกรมได้ให้ผู้ใช้ใส่ API Key ของ OpenRouteService โดยเข้าไปที่ลิงค์นี้ https://api.openrouteservice.org/ และทำการสมัครโดยสามารถหา API Key ได้ดังภาพด้านล่างนี้
![key](https://github.com/user-attachments/assets/48e41a09-fccb-46f4-a97e-4983b5d5779b)<br/>
### 2 Select Excel File:
- หลังจากใส่ API Key เรียบร้อยแล้ว ให้เราโหลดข้อมูลจากไฟล์ Excel .xlxs ซึ่งจะแสดงชื่อไฟล์ที่ทำการเลือกปัจจุบันด้านล่าง จากนั้นจะมีการการแสดงผลบนแผนที่เป็นไฟล์ HTML ซึ่งจะแสดงเส้นทางรถบรรทุกที่วางแผนไว้บนแผนที่มีการแสดงผลเส้นทางด้วยสี โดยแต่ละสีคือรถบรรทุกแต่ละคัน
![sel](https://github.com/user-attachments/assets/6fb2f01e-7da6-44d9-9186-0133bb856a01)<br/>
![key2](https://github.com/user-attachments/assets/74576b16-e713-443d-b99f-0bdb2a479fdb)<br/>
### 3 Generate HTML Map
- หลังจาก Select Excel ให้ทำการกดปุ่ม Generate HTML Map
![gen](https://github.com/user-attachments/assets/de03ffae-d6a1-4cd4-8e17-5c37fa8f6bc3)<br/>
![map](https://github.com/user-attachments/assets/1e26de10-ada6-4525-83d9-56a62417b530)
### 4 Export Excel: 
- หลังจาก Select Excel จะสามารถ Export ไฟล์เป็น excel โดยในแต่ละ sheet จะเป็นข้อมูล order ของรถบรรทุกแต่ละคัน
![export](https://github.com/user-attachments/assets/cab7b66c-8593-49b9-8d05-2cfbbced852a)<br/>
![exc](https://github.com/user-attachments/assets/39a8eb33-7c0d-4864-a889-e082b90805f2)<br />
### 5 CtkOptionMenu:
- หลังจากที่เรา Generate HTML Map เราจะสามารถเลือกกรองแค่เส้นทางของรถคันนั้น ๆ ได้ตามจำนวนรถทั้งหมด โดยหากทำการเลือก Truck 1 ก็จะแสดงผลเส้นทางเพียงแค่ Truck 1<br />
![tru](https://github.com/user-attachments/assets/93c007e6-6747-4bf4-a4f1-57e2f08313fb)<br />
![tru1](https://github.com/user-attachments/assets/1d07ba10-999f-4745-a0fc-7ad1c0b67e1a)
