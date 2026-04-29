# Car Booking System Overview

The **Car Booking System** is an internal web application designed to manage and streamline vehicle reservation requests, approvals, scheduling, and post-usage reporting within an organization.

It simulates a real-world logistics and transportation workflow, allowing users to submit booking requests, route them through a multi-level approval process, assign vehicles and drivers, and track operational and financial details after each trip.

The system is intended for:

* **Operational efficiency** by digitizing manual booking processes
* **Approval workflow management** across multiple roles
* **Fleet and driver coordination**
* **Financial tracking and reporting** for transportation costs
* **Audit and transparency** through timestamp tracking and logs

---

## Access and Usage

To access the Car Booking System:

1. Open: https://headcount.employeeskiosk.com
2. Login using email "benchakino04@gmail.com" password "123456"
3. Log in using the credentials
4. Use the system as an authenticated user

---

# Car Booking Process

## Before Appointment

### 1. Request Creation

The requester must fill out the Car Booking form with the following details:

* Requester Name
* Requester Department
* Date of Filing
* Client Name (e.g., Fiberhome)
* Company Assigned (e.g., Fiberhome – Primus)
* Number of Passengers
* Pickup Date
* Pickup Time
* Pickup Location
* Drop-off Location / Purpose

---

### 2. Approval Workflow

The request undergoes sequential approval from:

* Department Head
* RSB
* WVB

  * *Before approval, WVB assigns:*

    * Vehicle
    * Driver
* ODJ
* NAN

> The request may be **rejected** at any stage if requirements are not met.

---

### 3. Scheduling

Once approved, the booking is posted to the **Car Booking Schedule**.

---

## After Appointment

### 1. Trip Completion Encoding (RSB)

The following details are recorded:

* Time Returned
* Mileage (Start and End)
* Fuel Cost
* RFID Cost
* Parking Cost
* Carwash Cost
* Other Expenses

---

### 2. Validation

* Verified and signed by **RSB2**
* Acknowledged by **Requester**

---

# System Features

## Core Functionalities

* Dashboard with real-time booking details
* Timestamp tracking for every edit and approval
* Mobile-responsive dashboard access
* Printable and downloadable reports
* Export data to Excel

---

## Workflow Controls

* Booking can be **rejected by WVB** if no vehicle/driver is available
* Booking can be **cancelled/rejected even after approval** (emergency cases)
* Requester can edit booking **before RSB approval only**
* RSB and WVB have extended editing permissions

---

## Location-Based Billing

The system calculates charges based on destination:

* Mandaluyong, Pasig, San Juan, Makati, Taguig, Pasay → Php 500
* Manila, Quezon City, Malabon, Valenzuela, Parañaque → Php 1,000
* Bulacan, Cavite, Laguna, Rizal → Php 1,500
* Pampanga → Php 2,500
* Pangasinan, Quezon → Php 3,500
* Others → Custom pricing

---

## Driver Management

Drivers are mapped to companies:

* MPR → Primus
* DDC → Prime
* EAB → Prime
* CEP → Primus
* KMB → Primus
* WVB → Primus
* EHM → Primus
* Others → قابل configurable

---

## Vehicle Management

Vehicles are categorized by company:

* Innova Gray → Primetech
* Innova Red → Primus
* Innova Hybrid → Primus
* Tamaraw → Primus
* Others → Configurable

---

## Driver Fee Calculation

* Whole Day → Php 1,200
* Half Day → Php 600

---

## Financial and Reporting Features

* Automatic computation of:

  * Driver charges
  * Vehicle charges
* RFID usage and transaction reporting
* Searchable booking records
* Summary of charges per company

### Summary Approval Flow:

* Signed by RSB
* Signed by Payables
* Signed by MOM

---

# Key Capabilities

* Full audit trail via timestamps
* Role-based approvals and permissions
* End-to-end booking lifecycle tracking
* Financial transparency and reporting
* Mobile accessibility

---

# Notes

This system is designed to replicate real-world fleet management operations and enforce structured approval, accountability, and cost tracking across all transportation requests.
