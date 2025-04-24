# OPC and TCP Socket Integration System for Steel Pickling Process

**Discipline:** Distributed Systems for Automation (ELT011)  
**Professor:** Luiz T. S. Mendes — UFMG  
**Semester:** 2024/2  
**Language:** C/C++  
**IDE:** Visual Studio Community  

---

## 🧠 Project Overview

This project simulates the integration of a Windows-based OPC client with a Solaris-based process computer in a steel industry cold rolling mill. It aims to enable communication between a **PLCs-based OPC server** and a **process computer** using **TCP/IP sockets**. The main objective is to synchronize process variables and control setpoints for the steel pickling line.

---

## ⚙️ Architecture

The application consists of two primary modules:

1. **OPC Client Module**  
   - Reads process variables from a Matrikon OPC Simulation Server (asynchronously via callback).
   - Writes setpoints back to the server (synchronously).
   - Variables:
     - Acid concentration (T1/T2)
     - Temperature (T1/T2)
     - Strip speed

2. **TCP Client Module**  
   - Sends:
     - Periodic data messages (every 2 seconds)
     - On-demand setpoint requests (triggered by keyboard input `s`)
   - Receives:
     - ACKs and setpoint responses from the Solaris-based process computer
   - Reconnects automatically on communication loss.

---

## 🔗 Message Formats

All messages use ASCII strings separated by the symbol `$`.

### ➤ Process Data (sent to process computer)
NNNNN$555$C1$C2$T1$T2$SPD

swift
Copy
Edit
Example:  
`02736$555$4.5130$6.7244$054$088$02.01`

### ➤ ACK from process computer
NNNNN$000

shell
Copy
Edit

### ➤ Setpoint request (triggered by `s` key)
NNNNN$222

shell
Copy
Edit

### ➤ Setpoint response (from process computer)
NNNNN$100$T1_SP$T2_SP$C1_SP$C2_SP$SPD_SP

graphql
Copy
Edit

### ➤ ACK to setpoint response
NNNNN$999

yaml
Copy
Edit

---

## 🧩 OPC Simulation Items Mapping

| Variable | OPC Item (Matrikon) |
|----------|---------------------|
| C1 (%)   | `Random.Real4`      |
| C2 (%)   | `Random.Real8`      |
| T1 (°C)  | `Random.Int2`       |
| T2 (°C)  | `Random.Int4`       |
| SPD (cm/s)| `Triangle Waves.Real4` |
| SP T1    | `Bucket Brigade.Int2`   |
| SP T2    | `Bucket Brigade.Int4`   |
| SP C1    | `Bucket Brigade.Real4`  |
| SP C2    | `Bucket Brigade.Real8`  |
| SP SPD   | `Bucket Brigade.UInt1`  |

---

## 🖥 Console Output

All sent/received messages and OPC read/write operations are printed in the terminal. This allows full traceability of the process simulation and helps with debugging.

---

## 🧬 Recommended Architecture

Use a **multithreaded** approach:
- One thread for the OPC Client
- One thread for the TCP Client
- Shared memory or IPC with mutex/semaphore for inter-thread communication

Each module should remain functional even if the other fails temporarily.

---

## 📚 References

- Matrikon OPC Simulation Server: [www.matrikon.com](http://www.matrikon.com)
- OPC DA Overview: [OPC Foundation](https://opcfoundation.org/)
- Steel Pickling Processes: Usiminas, Rizzo (2022), ABM, etc.

---

> *This project was developed as part of the undergraduate course in Automation Engineering at UFMG.*
