# Truck Route Planner Version 2
แอปพลิเคชันนี้ช่วยให้ผู้ใช้สามารถเลือกไฟล์ Excel ที่มีข้อมูลคำสั่งซื้อ ประมวลผลข้อมูลเพื่อกำหนดเส้นทางที่มีประสิทธิภาพที่สุด และแสดงผลเส้นทางเหล่านี้บนแผนที่ นอกจากนี้ผู้ใช้ยังสามารถส่งออกเส้นทางที่วางแผนไว้ไปยังไฟล์ Excel ได้ ซึ่งเป็นการใช้ OpenRouteService (ORS) API มาใช้ในการคำนวณระยะทางและ ETA
โดยการใช้งานสามารถเข้าได้ด้วยไฟล์ exe ดังรูปนี้ <br />
## Update 2.2 !!!!
- เพิ่มการใส่เวลารถออก โดยผู้ใช้ต้องทำการใส่เวลาที่รถออกเป็น format เช่น 10:30 หรือ 11:30 เป็นต้น <br />
![1](https://github.com/user-attachments/assets/878ebe46-06db-451f-89e0-d17aa8cfdb27)<br />
- เพิ่ม Cummulative Distance และ Estimated time arrival (ETA) บน sidebar และในไฟล์ Excel ที่ Export<br />
![2](https://github.com/user-attachments/assets/eda89e50-da95-4aee-9ac1-a240974efd31)<br />
![icon](https://github.com/user-attachments/assets/151ccd08-ee57-4d4c-8076-f6ade0807c13)<br/>
![Screenshot 2024-07-27 152717](https://github.com/user-attachments/assets/dab84804-ec52-4d34-9e26-b62b3c4793ab)<br />
## การใช้งาน ต้องตามตามลำดับขึ้นดังนี้
### 1 OpenRouteService API Key
- การจะใช้งานโปรแกรมได้ให้ผู้ใช้ใส่ API Key ของ OpenRouteService โดยเข้าไปที่ลิงค์นี้ https://api.openrouteservice.org/ และทำการสมัครโดยสามารถหา API Key ได้ดังภาพด้านล่างนี้
![key](https://github.com/user-attachments/assets/48e41a09-fccb-46f4-a97e-4983b5d5779b)<br/>
### 2 Start time HH:MM
- ให้ทำการใส่เวลาที่รถเริ่มออก เพื่อการคำนวณระยะเวลากการเดินทางโดยประมาณในเส้นทาง โดยทำการใส่ ชั่วโมง:นาที เช่น 10:30 ดังภาพ
![time](https://github.com/user-attachments/assets/ad172a4c-405d-4b87-bbe7-7151a6f07b48)<br />
### 3 Select Excel File:
- หลังจากใส่ API Key เรียบร้อยแล้ว ให้เราโหลดข้อมูลจากไฟล์ Excel .xlxs ซึ่งจะแสดงชื่อไฟล์ที่ทำการเลือกปัจจุบันด้านล่าง จากนั้นจะมีการการแสดงผลบนแผนที่เป็นไฟล์ HTML ซึ่งจะแสดงเส้นทางรถบรรทุกที่วางแผนไว้บนแผนที่มีการแสดงผลเส้นทางด้วยสี โดยแต่ละสีคือรถบรรทุกแต่ละคัน
![sel](https://github.com/user-attachments/assets/6fb2f01e-7da6-44d9-9186-0133bb856a01)<br/>
![key2](https://github.com/user-attachments/assets/74576b16-e713-443d-b99f-0bdb2a479fdb)<br/>
### 4 Generate HTML Map
- หลังจาก Select Excel ให้ทำการกดปุ่ม Generate HTML Map
![4](https://github.com/user-attachments/assets/379e567e-46b6-4055-ae24-38826346eb3f)<br/>
### 5 Export Excel: 
- หลังจาก Select Excel จะสามารถ Export ไฟล์เป็น excel โดยในแต่ละ sheet จะเป็นข้อมูล order ของรถบรรทุกแต่ละคัน
![export](https://github.com/user-attachments/assets/cab7b66c-8593-49b9-8d05-2cfbbced852a)<br/>
![5](https://github.com/user-attachments/assets/b8848016-1cdd-4cda-8b12-6eddd3317b13)<br />
### 6 CtkOptionMenu:
- หลังจากที่เรา Generate HTML Map เราจะสามารถเลือกกรองแค่เส้นทางของรถคันนั้น ๆ ได้ตามจำนวนรถทั้งหมด โดยหากทำการเลือก Truck 1 ก็จะแสดงผลเส้นทางเพียงแค่ Truck 1<br />
![tru](https://github.com/user-attachments/assets/93c007e6-6747-4bf4-a4f1-57e2f08313fb)<br />
![6](https://github.com/user-attachments/assets/071a2efb-2f0f-408d-8502-277d79569097)

## ข้อจำกัด
- 1. หากลูกค้ามีการปักหมุด lat lng ที่ไม่ใกล้เคียงถนน จะทำให้โปรแกรมไม่สามารถหาเส้นทางที่ไปยังจุดนั้น ๆ ได้จะทำให้เกิด Bug ที่จะไม่แสดงเส้นทางและเวลาของรถบรรทุกคันที่บรรทุก order ชิ้นนั้น ๆ
- 2. OpenRouteService ยังไม่มีการนำเอาสภาพการจราจรเข้ามาเป็น Factor ในการคิดคำนวณเวลา ETA ทำให้มีโอกาสที่เวลาจะคลาดเคลื่อน 
- 3. ยังไม่สามารถจำกัดเวลาให้อยู่ในช่วงก่อน 19:00 น. ได้
