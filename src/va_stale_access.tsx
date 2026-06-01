import { useState, useMemo, useEffect, useCallback, useRef } from "react";
import * as XLSX from 'xlsx';

// ── Role normalisation ────────────────────────────────────────────────────────
const NORM_MAP: Record<string, string> = {
  "vaquery read only":          "va query read only",
  "quality reviewer":           "va quality reviewer",
  "va approval admin":          "va approvals admin",
  "va hr staff er/lr":          "va hr staff er-lr",
  "va hr supervisor er/lr":     "va hr supervisor er-lr",
  "va hr supervisor er-lr":     "va hr supervisor er-lr",
  "va background investigation":"va background investigator",
  "va hr recruit-place":        "va hr staff recruit-place",
  "va facility ssbp - chro":    "va facility ssbp - chro",
  "va manpower mgmt admin":     "va manpower mgmt admin",
  "va hr position mgmt pilot":  "va hr position mgmt pilot",
  "va hr position manager":     "va hr position manager",
  "va hris read only":          "va hris read only",
  "va ccoe supervisor":         "va ccoe supervisor",
  "va ccoe specialist":         "va ccoe specialist",
  "va hr handicap rno":         "va hr handicap rno",
  "va dcps hr":                 "va dcps hr",
  "va benefits manager":        "va benefits manager",
  "va hr staff tech review asst":"va hr staff tech review asst",
  "va worklife benefit specialist":"va worklife benefit specialist",
  "va hr staff benefits":       "va hr staff benefits",
  "va hr supervisor benefits":  "va hr supervisor benefits",
};
function normRole(r: string): string {
  const lo = r.trim().toLowerCase();
  return NORM_MAP[lo] || lo;
}
const BASE_NORM = new Set([
  "benefits ebss","va employee","va employee ess","va survey participant",
  "va hr base","txp redirect","va position update lite","va manager inflight",
  "va service chief",
]);
function isBase(r: string): boolean {
  const n = normRole(r);
  return BASE_NORM.has(n) || n.startsWith("va data region") ||
         n.startsWith("ssc mss") || n.startsWith("delegate");
}

// ── PD → required roles ───────────────────────────────────────────────────────
interface PDEntry {
  pdNum: string;
  position: string;
  grade: string;
  funcArea: string;
  roles: string[];
}

const PD_ROLE_DATA: PDEntry[] = [
  {pdNum:"99766S",position:"Deputy CHRO",             grade:"GS-201-14",funcArea:"HR Leadership",              roles:["VA HR Staff","VAQuery Read Only","VA Manager","VA Facility SSBP - CHRO"]},
  {pdNum:"99943S",position:"Assoc CHRO",               grade:"GS-201-14",funcArea:"HR Leadership",              roles:["VA HR Staff","VAQuery Read Only","VA Manager","VA Facility SSBP - CHRO"]},
  {pdNum:"99944S",position:"Assoc CHRO",               grade:"GS-201-14",funcArea:"HR Leadership",              roles:["VA HR Staff","VAQuery Read Only","VA Manager","VA Facility SSBP - CHRO"]},
  {pdNum:"99991S",position:"Strategic Business Partner",grade:"GS-201-14",funcArea:"HR Leadership",             roles:["VA HR Staff","VAQuery Read Only","VA Manager","VA Facility SSBP - CHRO"]},
  {pdNum:"99888S",position:"Supervisor SBU",           grade:"GS-201-13",funcArea:"HR Leadership",              roles:["VA HR Staff","VAQuery Read Only","VA Manager","VA HR Recruit-Place","VA HR Staff ER/LR","VA HR Supervisor Recruit-Place","VA HR Supervisor ER/LR"]},
  {pdNum:"99880S",position:"Lead HR Specialist SBU",   grade:"GS-201-13",funcArea:"HR Leadership",              roles:["VA HR Staff","VAQuery Read Only","VA HR Recruit-Place","VA HR Supervisor Recruit-Place"]},
  {pdNum:"99995S",position:"Supervisor",               grade:"GS-343-13",funcArea:"Manpower/Position Mgmt",     roles:["VA Manpower Mgmt Admin","VA HR Position Mgmt Pilot","VA HR Position Manager","VA Query Read Only","VA Manager"]},
  {pdNum:"99922S",position:"Mgmt & Program Analyst",   grade:"GS-343-12",funcArea:"Manpower/Position Mgmt",     roles:["VA Manpower Mgmt Admin","VA Query Read Only"]},
  {pdNum:"99640S",position:"Mgmt & Program Analyst",   grade:"GS-343-11",funcArea:"Manpower/Position Mgmt",     roles:["VA Manpower Mgmt Admin","VA Query Read Only"]},
  {pdNum:"99639S",position:"Mgmt & Program Analyst",   grade:"GS-343-09",funcArea:"Manpower/Position Mgmt",     roles:["VA Manpower Mgmt Admin","VA Query Read Only"]},
  {pdNum:"99638S",position:"Mgmt & Program Analyst",   grade:"GS-343-07",funcArea:"Manpower/Position Mgmt",     roles:["VA Manpower Mgmt Admin","VA Query Read Only"]},
  {pdNum:"99923S",position:"Position Manager",         grade:"GS-201-09",funcArea:"Manpower/Position Mgmt",     roles:["VA HR Position Mgmt Pilot","VA HR Position Manager","VA Query Read Only"]},
  {pdNum:"99691S",position:"Position Manager",         grade:"GS-201-07",funcArea:"Manpower/Position Mgmt",     roles:["VA HR Position Mgmt Pilot","VA HR Position Manager","VA Query Read Only"]},
  {pdNum:"99986S",position:"Supervisor",               grade:"GS-201-13",funcArea:"Classification",             roles:["VA HR Staff","VAQuery Read Only","VA Manager"]},
  {pdNum:"99985S",position:"Lead HR Specialist",       grade:"GS-201-13",funcArea:"Classification",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99984S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"Classification",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99674S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"Classification",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99673S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"Classification",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99672S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"Classification",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99981S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"Classification",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99670S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"Classification",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99669S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"Classification",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99992S",position:"Supervisor",               grade:"GS-201-13",funcArea:"Compensation",               roles:["VA HR Staff","VAQuery Read Only","VA Manager","VA CCOE Supervisor","VA HR Staff Recruit-Place","VA HR Supervisor Recruit-Place"]},
  {pdNum:"99982S",position:"Lead HR Specialist",       grade:"GS-201-13",funcArea:"Compensation",               roles:["VA HR Staff","VAQuery Read Only","VA CCOE Supervisor","VA HR Staff Recruit-Place","VA HR Supervisor Recruit-Place"]},
  {pdNum:"99993S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"Compensation",               roles:["VA HR Staff","VAQuery Read Only","VA CCOE Specialist","VA HR Staff Recruit-Place"]},
  {pdNum:"99682S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"Compensation",               roles:["VA HR Staff","VAQuery Read Only","VA CCOE Specialist","VA HR Staff Recruit-Place"]},
  {pdNum:"99681S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"Compensation",               roles:["VA HR Staff","VAQuery Read Only","VA CCOE Specialist","VA HR Staff Recruit-Place"]},
  {pdNum:"99680S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"Compensation",               roles:["VA HR Staff","VAQuery Read Only","VA CCOE Specialist","VA HR Staff Recruit-Place"]},
  {pdNum:"99987S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"Compensation",               roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99676S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"Compensation",               roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99675S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"Compensation",               roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99994S",position:"Supervisor",               grade:"GS-201-13",funcArea:"HRIS",                       roles:["VA HR Staff","VA Approval Admin","VAQuery Read Only","VA Manager","VA HRIS Read Only"]},
  {pdNum:"99369S",position:"Lead HR Specialist",       grade:"GS-201-13",funcArea:"HRIS",                       roles:["VA HR Staff","VA Approval Admin","VAQuery Read Only","VA HRIS Read Only"]},
  {pdNum:"99997S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"HRIS",                       roles:["VA HR Staff","VA Approval Admin","VAQuery Read Only","VA HRIS Read Only"]},
  {pdNum:"99687S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"HRIS",                       roles:["VA HR Staff","VA Approval Admin","VAQuery Read Only","VA HRIS Read Only"]},
  {pdNum:"99686S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"HRIS",                       roles:["VA HR Staff","VA Approval Admin","VAQuery Read Only","VA HRIS Read Only"]},
  {pdNum:"99685S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"HRIS",                       roles:["VA HR Staff","VA Approval Admin","VAQuery Read Only","VA HRIS Read Only"]},
  {pdNum:"99370S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"HRIS",                       roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99365S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"HRIS",                       roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99366S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"HRIS",                       roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99813S",position:"Supervisor",               grade:"GS-201-13",funcArea:"Technical Review",           roles:["Quality Reviewer","VA HR Handicap RNO","VA DCPS HR","VA Query Read Only","VA Benefits Manager","VA Manager"]},
  {pdNum:"99852S",position:"Lead HR Specialist",       grade:"GS-201-13",funcArea:"Technical Review",           roles:["Quality Reviewer","VA HR Handicap RNO","VA DCPS HR","VA Benefits Manager","VA Query Read Only"]},
  {pdNum:"99851S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"Technical Review",           roles:["Quality Reviewer","VA HR Handicap RNO","VA DCPS HR","VA Benefits Manager","VA Query Read Only"]},
  {pdNum:"99637S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"Technical Review",           roles:["Quality Reviewer","VA HR Handicap RNO","VA DCPS HR","VA Benefits Manager","VA Query Read Only"]},
  {pdNum:"99636S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"Technical Review",           roles:["Quality Reviewer","VA HR Handicap RNO","VA DCPS HR","VA Benefits Manager","VA Query Read Only"]},
  {pdNum:"99635S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"Technical Review",           roles:["Quality Reviewer","VA HR Handicap RNO","VA DCPS HR","VA Benefits Manager","VA Query Read Only"]},
  {pdNum:"99609S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"Technical Review",           roles:["VA HR Staff","VA HR Staff Tech Review Asst","VA Query Read Only"]},
  {pdNum:"99608S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"Technical Review",           roles:["VA HR Staff","VA HR Staff Tech Review Asst","VA Query Read Only"]},
  {pdNum:"99607S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"Technical Review",           roles:["VA HR Staff","VA HR Staff Tech Review Asst","VA Query Read Only"]},
  {pdNum:"99971S",position:"Supervisor",               grade:"GS-201-13",funcArea:"Quality Assurance",          roles:["VA HR Staff","VAQuery Read Only","VA Manager"]},
  {pdNum:"99970S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"Quality Assurance",          roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99533S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"Quality Assurance",          roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99534S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"Quality Assurance",          roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99535S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"Quality Assurance",          roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99938S",position:"Supervisor",               grade:"GS-201-13",funcArea:"HR Development",             roles:["VA HR Staff","VAQuery Read Only","VA Manager"]},
  {pdNum:"99939S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"HR Development",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99648S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"HR Development",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99647S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"HR Development",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99646S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"HR Development",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99937S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"HR Development",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99536S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"HR Development",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99532S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"HR Development",             roles:["VA HR Staff","VA Query Read Only"]},
  {pdNum:"99972S",position:"Deputy SSU Supervisor",    grade:"GS-201-14",funcArea:"Recruitment and Placement",  roles:["VA HR Supervisor Recruit-Place","VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA Manager","VA HR Recruiter"]},
  {pdNum:"99973S",position:"Supervisor",               grade:"GS-201-13",funcArea:"Recruitment and Placement",  roles:["VA HR Supervisor Recruit-Place","VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA Manager","VA HR Recruiter"]},
  {pdNum:"99974S",position:"Lead HR Specialist",       grade:"GS-201-13",funcArea:"Recruitment and Placement",  roles:["VA HR Supervisor Recruit-Place","VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA HR Recruiter"]},
  {pdNum:"99989S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"Recruitment and Placement",  roles:["VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA HR Recruiter"]},
  {pdNum:"99678S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"Recruitment and Placement",  roles:["VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA HR Recruiter"]},
  {pdNum:"99677S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"Recruitment and Placement",  roles:["VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA HR Recruiter"]},
  {pdNum:"99679S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"Recruitment and Placement",  roles:["VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA HR Recruiter"]},
  {pdNum:"99721S",position:"HR Specialist – Physician Recruiter",grade:"GS-201-12",funcArea:"Recruitment and Placement",roles:["VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA HR Recruiter"]},
  {pdNum:"99690S",position:"HR Specialist – Physician Recruiter",grade:"GS-201-11",funcArea:"Recruitment and Placement",roles:["VA HR Staff Recruit-Place","VA HR Staff","VAQuery Read Only","VA HR Recruiter"]},
  {pdNum:"99947S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"Recruitment and Placement",  roles:["VA HR Staff Recruit-Place","VA HR Staff","VA Query Read Only"]},
  {pdNum:"99650S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"Recruitment and Placement",  roles:["VA HR Staff Recruit-Place","VA HR Staff","VA Query Read Only"]},
  {pdNum:"99649S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"Recruitment and Placement",  roles:["VA HR Staff Recruit-Place","VA HR Staff","VA Query Read Only"]},
  {pdNum:"99979S",position:"Deputy SSU Supervisor",    grade:"GS-201-14",funcArea:"ER/LR",                      roles:["VA HR Supervisor ER-LR","VA HR Staff ER-LR","VA HR Staff","VAQuery Read Only","VA Manager","VA Worklife Benefit Specialist"]},
  {pdNum:"99980S",position:"Supervisor",               grade:"GS-201-13",funcArea:"ER/LR",                      roles:["VA HR Supervisor ER-LR","VA HR Staff ER-LR","VA HR Staff","VAQuery Read Only","VA Manager","VA Worklife Benefit Specialist"]},
  {pdNum:"99969S",position:"Lead HR Specialist",       grade:"GS-201-13",funcArea:"ER/LR",                      roles:["VA HR Supervisor ER-LR","VA HR Staff ER-LR","VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist"]},
  {pdNum:"99976S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"ER/LR",                      roles:["VA HR Staff ER-LR","VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist"]},
  {pdNum:"99666S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"ER/LR",                      roles:["VA HR Staff ER-LR","VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist"]},
  {pdNum:"99665S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"ER/LR",                      roles:["VA HR Staff ER-LR","VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist"]},
  {pdNum:"99664S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"ER/LR",                      roles:["VA HR Staff ER-LR","VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist"]},
  {pdNum:"99978S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"ER/LR",                      roles:["VA HR Staff ER-LR","VA HR Staff","VA Worklife Benefit Specialist","VA Query Read Only"]},
  {pdNum:"99668S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"ER/LR",                      roles:["VA HR Staff ER-LR","VA HR Staff","VA Worklife Benefit Specialist","VA Query Read Only"]},
  {pdNum:"99667S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"ER/LR",                      roles:["VA HR Staff ER-LR","VA HR Staff","VA Worklife Benefit Specialist","VA Query Read Only"]},
  {pdNum:"99934S",position:"Supervisor",               grade:"GS-201-13",funcArea:"Work Life Benefits",         roles:["VA HR Supervisor Benefits","VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VAQuery Read Only","VA Manager"]},
  {pdNum:"99934S",position:"Supervisor",               grade:"GS-201-13",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VAQuery Read Only","VA Manager","VA Worklife Benefit Specialist","VA HR Supervisor Benefits","VA HR Staff Benefits"]},
  {pdNum:"99899S",position:"Lead HR Specialist",       grade:"GS-201-12",funcArea:"Work Life Benefits",         roles:["VA HR Supervisor Benefits","VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VA Query Read Only"]},
  {pdNum:"99935S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"Work Life Benefits",         roles:["VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VA Query Read Only"]},
  {pdNum:"99643S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"Work Life Benefits",         roles:["VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VA Query Read Only"]},
  {pdNum:"99642S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"Work Life Benefits",         roles:["VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VA Query Read Only"]},
  {pdNum:"99641S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"Work Life Benefits",         roles:["VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VA Query Read Only"]},
  {pdNum:"99936S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"Work Life Benefits",         roles:["VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VA Query Read Only"]},
  {pdNum:"99936S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99645S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"Work Life Benefits",         roles:["VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VA Query Read Only"]},
  {pdNum:"99645S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99644S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"Work Life Benefits",         roles:["VA HR Staff Benefits","VA HR Staff","VA Worklife Benefit Specialist","VA DCPS HR","VA Query Read Only"]},
  {pdNum:"99644S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99977S",position:"Lead HR Specialist",       grade:"GS-201-12",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist","VA HR Supervisor ER/LR","VA HR Staff ER/LR"]},
  {pdNum:"99975S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist","VA HR Staff ER/LR"]},
  {pdNum:"99663S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist","VA HR Staff ER/LR"]},
  {pdNum:"99662S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist","VA HR Staff ER/LR"]},
  {pdNum:"99661S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"Reasonable Accommodations",  roles:["VA HR Staff","VA Query Read Only","VA Worklife Benefit Specialist","VA HR Staff ER/LR"]},
  {pdNum:"99961S",position:"Supervisor",               grade:"GS-201-13",funcArea:"Worker's Compensation",      roles:["VA HR Staff","VAQuery Read Only","VA Manager","VA HR Supervisor Benefits","VA HR Staff Benefits"]},
  {pdNum:"99959S",position:"Lead HR Specialist",       grade:"GS-0201-13",funcArea:"Worker's Compensation",     roles:["VA HR Staff","VA Query Read Only","VA HR Supervisor Benefits","VA HR Staff Benefits"]},
  {pdNum:"99958S",position:"HR Specialist",            grade:"GS-201-12",funcArea:"Worker's Compensation",      roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99657S",position:"HR Specialist",            grade:"GS-201-11",funcArea:"Worker's Compensation",      roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99656S",position:"HR Specialist",            grade:"GS-201-09",funcArea:"Worker's Compensation",      roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99655S",position:"HR Specialist",            grade:"GS-201-07",funcArea:"Worker's Compensation",      roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99660S",position:"HR Assistant",             grade:"GS-203-07",funcArea:"Worker's Compensation",      roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99659S",position:"HR Assistant",             grade:"GS-203-06",funcArea:"Worker's Compensation",      roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99658S",position:"HR Assistant",             grade:"GS-203-05",funcArea:"Worker's Compensation",      roles:["VA HR Staff","VAQuery Read Only","VA HR Staff Benefits"]},
  {pdNum:"99948S",position:"VISN Personnel Security Officer",grade:"GS-0080-13",funcArea:"Personnel Security",  roles:["VA Background Investigation","VA Query Read Only","VA Manager"]},
  {pdNum:"99949S",position:"Supervisor",               grade:"GS-0080-12",funcArea:"Personnel Security",        roles:["VA Background Investigation","VA Query Read Only","VA Manager"]},
  {pdNum:"99953S",position:"Lead Personnel Security",  grade:"GS-0080-11",funcArea:"Personnel Security",        roles:["VA Background Investigation","VA Query Read Only"]},
  {pdNum:"99950S",position:"Suitability Specialist",   grade:"GS-0080-11",funcArea:"Personnel Security",        roles:["VA Background Investigation","VA Query Read Only"]},
  {pdNum:"99652S",position:"Suitability Specialist",   grade:"GS-0080-09",funcArea:"Personnel Security",        roles:["VA Background Investigation","VA Query Read Only"]},
  {pdNum:"99651S",position:"Suitability Specialist",   grade:"GS-0080-07",funcArea:"Personnel Security",        roles:["VA Background Investigation","VA Query Read Only"]},
  {pdNum:"99952S",position:"Personnel Security Asst",  grade:"GS-0086-07",funcArea:"Personnel Security",        roles:["VA Query Read Only"]},
  {pdNum:"99654S",position:"Personnel Security Asst",  grade:"GS-0086-06",funcArea:"Personnel Security",        roles:["VA Query Read Only"]},
  {pdNum:"99653S",position:"Personnel Security Asst",  grade:"GS-0086-05",funcArea:"Personnel Security",        roles:["VA Query Read Only"]},
];

const FUNC_AREA_COLORS: Record<string, string> = {
  "HR Leadership":"#5B3FA0","Manpower/Position Mgmt":"#185FA5","Classification":"#0C7A52",
  "Compensation":"#854F0B","HRIS":"#A32D2D","Technical Review":"#2D5FA3",
  "Quality Assurance":"#5A2DA3","HR Development":"#1A7A6E","Recruitment and Placement":"#7A501A",
  "ER/LR":"#6B2A6B","Work Life Benefits":"#2A6B3A","Reasonable Accommodations":"#3A2A6B",
  "Worker's Compensation":"#6B3A2A","Personnel Security":"#2A4A6B"
};

// ── Stale / mismatch data ────────────────────────────────────────────────────
interface VaDataRole {
  role: string;
  type: string;
  mismatch: boolean;
  reason?: string;
}

interface UserRecord {
  emplId: string;
  name: string;
  userId: string;
  email: string;
  station: string;
  loc: string;
  title: string;
  orgCode: string;
  poid: string;
  isLocked: number;
  lastLogin: string | null;
  vaDataRoles: VaDataRole[];
}

interface ExceptionRecord {
  note: string;
  expiry: string;
  approver: string;
  savedAt: string;
}

const STALE_DATA: UserRecord[] = [{"emplId":"00341239","name":"Jons,Denise A","userId":"denise.jons","email":"denise.jons@va.gov","station":"589","loc":"VAMC Kansas City MO","title":"Supervisory Clinical Lab Scientist","orgCode":"2070","poid":"1253","isLocked":1,"lastLogin":"2025-11-19","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2070","type":"Org Cd","mismatch":false}]},{"emplId":"00468511","name":"Higgins,Cheryl R","userId":"cheryl.higgins","email":"cheryl.higgins@va.gov","station":"589","loc":"VAMC Kansas City MO","title":"Inventory Management Specialist","orgCode":"2595","poid":"1253","isLocked":1,"lastLogin":"2025-04-24","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2595","type":"Org Cd","mismatch":false}]},{"emplId":"00287399","name":"May,Chad W","userId":"chad.may","email":"chad.may@va.gov","station":"589","loc":"Colmery-O'Neil VAMC Topeka KS","title":"Voluntary Services Assistant","orgCode":"2310","poid":"1479","isLocked":1,"lastLogin":"2025-06-27","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2518","type":"Org Cd","mismatch":true,"reason":"Role Org Cd 2518 ≠ user Org Code 2310"}]},{"emplId":"00493901","name":"Anderson,Dustin R","userId":"Dustin.Anderson1","email":"Dustin.Anderson1@va.gov","station":"589","loc":"Colmery-O'Neil VAMC Topeka KS","title":"Program Support Assistant (OA)","orgCode":"2516","poid":"1479","isLocked":1,"lastLogin":"2025-06-26","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2516","type":"Org Cd","mismatch":false}]},{"emplId":"00490388","name":"Everett,Matthew C","userId":"matthew.everett","email":"matthew.everett@va.gov","station":"589","loc":"Dwight D. Eisenhower DVAMC","title":"Administrative Officer","orgCode":"2560","poid":"1479","isLocked":1,"lastLogin":"2025-09-04","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2560","type":"Org Cd","mismatch":false}]},{"emplId":"00341624","name":"Stephens,Suzanne H","userId":"suzanne.stephens","email":"suzanne.stephens@va.gov","station":"657","loc":"MC/VCS/NCA HCM Saint Louis MO","title":"Purchasing Agent","orgCode":"2088","poid":"1783","isLocked":1,"lastLogin":"2025-04-01","vaDataRoles":[{"role":"VA Data Stn 657 Org Cd 2050","type":"Org Cd","mismatch":true,"reason":"Role Org Cd 2050 ≠ user Org Code 2088"}]},{"emplId":"00345988","name":"Sutton,Sara E","userId":"sara.sutton","email":"sara.sutton@va.gov","station":"657","loc":"MC/VCS/NCA HCM Saint Louis MO","title":"Logistics Management Specialist","orgCode":"2595","poid":"1783","isLocked":0,"lastLogin":null,"vaDataRoles":[{"role":"VA Data Stn 657 Org Cd 2595","type":"Org Cd","mismatch":false}]},{"emplId":"00337459","name":"Pruitt,Michelle M","userId":"michelle.pruitt2","email":"michelle.pruitt2@va.gov","station":"589","loc":"Harry S. Truman Memorial VH","title":"Health System Specialist","orgCode":"2323","poid":"4046","isLocked":1,"lastLogin":"2025-04-21","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2079","type":"Org Cd","mismatch":true,"reason":"Role Org Cd 2079 ≠ user Org Code 2323"}]},{"emplId":"00338402","name":"Hosenfelt,Karen M","userId":"karen.clark1","email":"KAREN.HOSENFELT@VA.GOV","station":"589","loc":"Harry S. Truman Memorial VH","title":"Program Analyst (Informatics)","orgCode":"2323","poid":"4046","isLocked":1,"lastLogin":"2025-11-25","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2000","type":"Org Cd","mismatch":true,"reason":"Role Org Cd 2000 ≠ user Org Code 2323"}]}];

const MISMATCH_DATA: UserRecord[] = [{"emplId":"00642794","name":"Jones,Tracy L","userId":"Tracy.Jones9","email":"Tracy.Jones9@va.gov","station":"657","loc":"John J. Pershing VAMC","title":"Administrative Officer","orgCode":"2033","poid":"1128","isLocked":0,"lastLogin":"2026-05-13","vaDataRoles":[{"role":"VA Data Stn 657 Org Cd 2026","type":"Org Cd","mismatch":true,"reason":"Role Org Cd 2026 ≠ user Org Code 2033"},{"role":"VA Data Stn 657 Org Cd 2033","type":"Org Cd","mismatch":false},{"role":"VA Data Stn 657 Org Cd 2115","type":"Org Cd","mismatch":true},{"role":"VA Data Stn 657 Org Cd 2260","type":"Org Cd","mismatch":true},{"role":"VA Data Stn 657 Org Cd 2284","type":"Org Cd","mismatch":true},{"role":"VA Data Stn 657 Org Cd 3327","type":"Org Cd","mismatch":true}]},{"emplId":"00332553","name":"Deford,Mary B","userId":"mary.deford","email":"mary.deford@va.gov","station":"657","loc":"John J. Pershing VAMC","title":"Medical Support Assistant (Advanced)","orgCode":"2448","poid":"1128","isLocked":1,"lastLogin":"2026-01-21","vaDataRoles":[{"role":"VA Data Stn 657 Org Cd 2115","type":"Org Cd","mismatch":true}]},{"emplId":"00546910","name":"Thurman,Suzanne C","userId":"suzanna.thurman","email":"Suzanne.Thurman@va.gov","station":"657","loc":"John J. Pershing VAMC","title":"Program Analyst (Informatics)","orgCode":"2448","poid":"1128","isLocked":0,"lastLogin":"2026-04-07","vaDataRoles":[{"role":"VA Data Stn 657 Org Cd 2500","type":"Org Cd","mismatch":true}]},{"emplId":"00554721","name":"Kimsey,David S","userId":"david.kimsey","email":"david.kimsey@va.gov","station":"657","loc":"VA Clinic Cape Girardeau MO","title":"Supervisory Health System Specialist","orgCode":"2186","poid":"1128","isLocked":0,"lastLogin":"2026-05-28","vaDataRoles":[{"role":"VA Data Stn 657 Org Cd 2500","type":"Org Cd","mismatch":true}]},{"emplId":"00665681","name":"McMath,Justin M","userId":"justin.mcmath","email":"justin.mcmath@va.gov","station":"589","loc":"VAMC Kansas City MO","title":"Administrative Officer","orgCode":"3327","poid":"1253","isLocked":0,"lastLogin":"2026-05-15","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2147","type":"Org Cd","mismatch":true},{"role":"VA Data Stn 589 Org Cd 3327","type":"Org Cd","mismatch":false}]},{"emplId":"00362870","name":"Streitwieser,Timothy C","userId":"timothy.streitwieser","email":"Timothy.Streitwieser@va.gov","station":"589","loc":"VAMC Kansas City MO","title":"SENIOR PSYCHOLOGIST PROGRAM MANAGER","orgCode":"2135","poid":"1253","isLocked":0,"lastLogin":"2026-05-27","vaDataRoles":[{"role":"VA Data Stn 589 Org Cd 2173","type":"Org Cd","mismatch":true}]}];

const TOTAL_ADMIN = 298;
const STORAGE_KEY = "va_access_exceptions";
const TODAY = new Date("2026-05-29");

function isExpired(ex: ExceptionRecord | undefined): boolean {
  return ex?.expiry ? new Date(ex.expiry) < TODAY : false;
}
function isActive(ex: ExceptionRecord | undefined): boolean {
  return !!ex && !isExpired(ex);
}

// ── Shared small components ───────────────────────────────────────────────────
function RoleBadge({ r }: { r: VaDataRole }) {
  return (
    <span title={r.reason || ""} style={{display:"inline-flex",alignItems:"center",gap:3,fontSize:11,padding:"2px 7px",borderRadius:20,margin:"2px 2px 2px 0",background:r.mismatch?"#fff0e6":"#E6F1FB",color:r.mismatch?"#854F0B":"#0C447C"}}>
      {r.mismatch && <span style={{fontWeight:700}}>!</span>}{r.role}
    </span>
  );
}

function StatusPill({ status }: { status: string }) {
  const cfg: Record<string, {bg:string;color:string;label:string}> = {
    compliant: {bg:"#e6f7ee",color:"#0C7A52",label:"✓ Compliant"},
    missing:   {bg:"#fff0e6",color:"#A32D2D",label:"✗ Missing roles"},
    partial:   {bg:"#fff8e6",color:"#854F0B",label:"~ Partial match"},
  };
  const c = cfg[status] || {bg:"#f0f0f0",color:"#555",label:status};
  return <span style={{fontSize:11,padding:"2px 9px",borderRadius:20,background:c.bg,color:c.color,fontWeight:500,whiteSpace:"nowrap"}}>{c.label}</span>;
}

// ── Exception form ────────────────────────────────────────────────────────────
interface ExceptionFormValues {
  note: string;
  expiry: string;
  approver: string;
}

const EMPTY_FORM: ExceptionFormValues = {note:"",expiry:"",approver:""};

function ExceptionForm({
  d, ex, onSave, onRemove, onCancel, saving, colSpan = 9
}: {
  d: UserRecord;
  ex: ExceptionRecord | undefined;
  onSave: (id: string, f: ExceptionFormValues) => void;
  onRemove: (id: string) => void;
  onCancel: () => void;
  saving: boolean;
  colSpan?: number;
}) {
  const [form, setForm] = useState<ExceptionFormValues>(
    ex ? {note:ex.note||"",expiry:ex.expiry||"",approver:ex.approver||""} : EMPTY_FORM
  );
  const valid = form.note.trim() && form.expiry && form.approver.trim();
  const inp: React.CSSProperties = {padding:"6px 9px",fontSize:13,border:"1px solid #ddd",borderRadius:6,width:"100%",boxSizing:"border-box"};
  const lbl: React.CSSProperties = {fontSize:12,color:"#666",display:"block",marginBottom:3,fontWeight:500};
  return (
    <tr style={{background:"#f7faff",borderBottom:"1px solid #e0eaff"}}>
      <td colSpan={colSpan} style={{padding:"12px 16px"}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:12}}>
          <div style={{gridColumn:"1 / -1"}}>
            <label style={lbl}>Exception note <span style={{color:"#A32D2D"}}>*</span></label>
            <textarea value={form.note} onChange={e=>setForm(f=>({...f,note:e.target.value}))} rows={2} placeholder="Reason for exception" style={{...inp,resize:"vertical"}}/>
          </div>
          <div>
            <label style={lbl}>Approved by <span style={{color:"#A32D2D"}}>*</span></label>
            <input value={form.approver} onChange={e=>setForm(f=>({...f,approver:e.target.value}))} placeholder="Name or user ID" style={inp}/>
          </div>
          <div>
            <label style={lbl}>Expiration date <span style={{color:"#A32D2D"}}>*</span></label>
            <input type="date" value={form.expiry} onChange={e=>setForm(f=>({...f,expiry:e.target.value}))} min="2026-05-29" style={inp}/>
          </div>
          <div style={{display:"flex",alignItems:"flex-end",gap:8}}>
            <button onClick={()=>onSave(d.emplId,form)} disabled={saving||!valid} style={{padding:"6px 16px",fontSize:13,border:"none",borderRadius:6,cursor:"pointer",background:"#185FA5",color:"white",opacity:valid?1:0.5}}>{saving?"Saving…":"Save"}</button>
            <button onClick={onCancel} style={{padding:"6px 12px",fontSize:13,border:"1px solid #ddd",borderRadius:6,cursor:"pointer",background:"white",color:"#555"}}>Cancel</button>
            {ex && <button onClick={()=>onRemove(d.emplId)} style={{padding:"6px 12px",fontSize:13,border:"1px solid #f0c0c0",borderRadius:6,cursor:"pointer",background:"white",color:"#A32D2D"}}>Remove</button>}
          </div>
        </div>
        {ex && <div style={{fontSize:12,color:"#888"}}>Last saved: {ex.savedAt} · Approved by: {ex.approver}</div>}
      </td>
    </tr>
  );
}

// ── VA Data / Stale table ─────────────────────────────────────────────────────
function DataTable({
  data, exceptions, onSave, onRemove, saving, showMismatchCount
}: {
  data: UserRecord[];
  exceptions: Record<string, ExceptionRecord>;
  onSave: (id: string, f: ExceptionFormValues) => void;
  onRemove: (id: string) => void;
  saving: boolean;
  showMismatchCount: boolean;
}) {
  const [search, setSearch] = useState("");
  const [station, setStation] = useState("");
  const [showExcepted, setShowExcepted] = useState(true);
  const [openForm, setOpenForm] = useState<string | null>(null);

  const stations = useMemo(() => [...new Set(data.map(d => d.station))].sort(), [data]);
  const exceptedCount = data.filter(d => isActive(exceptions[d.emplId])).length;
  const expiredCount = data.filter(d => exceptions[d.emplId] && isExpired(exceptions[d.emplId])).length;

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return data.filter(d => {
      if (!showExcepted && isActive(exceptions[d.emplId])) return false;
      const mQ = !q || [d.name,d.emplId,d.email,d.station,d.loc,d.title,d.orgCode,d.poid].some(v => v && String(v).toLowerCase().includes(q));
      return mQ && (!station || d.station === station);
    });
  }, [search,station,showExcepted,exceptions,data]);

  const th: React.CSSProperties = {padding:"8px 10px",textAlign:"left",fontWeight:500,fontSize:12,color:"#666",borderBottom:"1px solid #e5e5e5",whiteSpace:"nowrap"};
  const td: React.CSSProperties = {padding:"8px 10px",verticalAlign:"top"};

  return (
    <>
      {expiredCount > 0 && (
        <div style={{background:"#fff0f0",border:"1px solid #f0c0c0",borderRadius:8,padding:"8px 14px",fontSize:13,marginBottom:12,color:"#A32D2D"}}>
          ⚠ {expiredCount} exception{expiredCount > 1 ? "s have" : " has"} expired.
        </div>
      )}
      <div style={{display:"flex",gap:8,marginBottom:"1rem",flexWrap:"wrap",alignItems:"center"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name, ID, email, org code…" style={{flex:1,minWidth:180,padding:"6px 10px",fontSize:13,border:"1px solid #ddd",borderRadius:6}}/>
        <select value={station} onChange={e=>setStation(e.target.value)} style={{padding:"6px 10px",fontSize:13,border:"1px solid #ddd",borderRadius:6}}>
          <option value="">All stations</option>
          {stations.map(s => <option key={s} value={s}>Station {s}</option>)}
        </select>
        <button onClick={()=>setShowExcepted(v=>!v)} style={{padding:"6px 14px",fontSize:13,border:"1px solid #ddd",borderRadius:6,cursor:"pointer",background:showExcepted?"#E6F1FB":"#fff",color:showExcepted?"#0C447C":"#555"}}>
          {showExcepted ? "Hide excepted" : `Show excepted (${exceptedCount})`}
        </button>
      </div>
      <div style={{border:"1px solid #e5e5e5",borderRadius:10,overflow:"hidden"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead><tr style={{background:"#f9f9f9"}}>
            <th style={th}>Name / Email</th>
            <th style={th}>Empl ID</th>
            <th style={th}>Position title</th>
            <th style={th}>Org Code</th>
            {showMismatchCount && <th style={th}>POID</th>}
            <th style={th}>Station</th>
            <th style={th}>Last login</th>
            <th style={th}>VA Data roles</th>
            <th style={th}>Exception</th>
          </tr></thead>
          <tbody>
            {filtered.length === 0
              ? <tr><td colSpan={showMismatchCount?9:8} style={{textAlign:"center",padding:"2rem",color:"#999",fontSize:13}}>No records match.</td></tr>
              : filtered.map(d => {
                  const ex = exceptions[d.emplId];
                  const active = isActive(ex), expired = ex && isExpired(ex);
                  const mCount = d.vaDataRoles.filter(r => r.mismatch).length;
                  return [
                    <tr key={d.emplId} style={{borderBottom:openForm===d.emplId?"none":"1px solid #f0f0f0",background:active?"#f0f6ff":expired?"#fff8f0":"white"}}>
                      <td style={td}><div style={{fontWeight:500}}>{d.name}</div><div style={{fontSize:11,color:"#888"}}>{d.email}</div></td>
                      <td style={td}>{d.emplId}</td>
                      <td style={td}>{d.title}</td>
                      <td style={td}><span style={{fontWeight:500}}>{d.orgCode}</span></td>
                      {showMismatchCount && <td style={td}>{d.poid}</td>}
                      <td style={td}><div style={{fontWeight:500}}>{d.station}</div><div style={{fontSize:11,color:"#888"}}>{d.loc}</div></td>
                      <td style={{...td,fontWeight:500,color:d.lastLogin?"#854F0B":"#A32D2D"}}>{d.lastLogin||"Never"}</td>
                      <td style={td}>
                        {d.vaDataRoles.map((r,i) => <RoleBadge key={i} r={r}/>)}
                        {showMismatchCount && mCount > 0 && <div style={{fontSize:11,color:"#854F0B",marginTop:4}}>{mCount} mismatch{mCount!==1?"es":""}</div>}
                      </td>
                      <td style={td}>
                        {active && <span style={{display:"inline-block",fontSize:11,padding:"2px 8px",borderRadius:20,background:"#E6F1FB",color:"#0C447C",marginBottom:4}}>Excepted · {ex!.expiry}</span>}
                        {expired && <span style={{display:"inline-block",fontSize:11,padding:"2px 8px",borderRadius:20,background:"#fff0e6",color:"#854F0B",marginBottom:4}}>Expired {ex!.expiry}</span>}
                        <div>
                          <button onClick={()=>setOpenForm(openForm===d.emplId?null:d.emplId)} style={{fontSize:12,padding:"3px 10px",border:"1px solid #ddd",borderRadius:6,cursor:"pointer",background:"white",color:"#333"}}>
                            {active||expired?"Edit":"Add exception"}
                          </button>
                        </div>
                      </td>
                    </tr>,
                    openForm === d.emplId && (
                      <ExceptionForm key={d.emplId+"_f"} d={d} ex={ex} onSave={(id,f)=>{onSave(id,f);setOpenForm(null);}} onRemove={id=>{onRemove(id);setOpenForm(null);}} onCancel={()=>setOpenForm(null)} saving={saving} colSpan={showMismatchCount?9:8}/>
                    )
                  ];
                })
            }
          </tbody>
        </table>
      </div>
      <p style={{fontSize:12,color:"#999",marginTop:8,textAlign:"right"}}>
        Showing {filtered.length} of {data.length} flagged · System: HR-Smart · Role: VA Admin Officer
      </p>
    </>
  );
}

// ── HR Role Lookup ────────────────────────────────────────────────────────────
function HRRoleLookup() {
  const [query, setQuery] = useState("");
  const [filterArea, setFilterArea] = useState("");
  const [mode, setMode] = useState<"search"|"browse">("search");

  const funcAreas = useMemo(() => [...new Set(PD_ROLE_DATA.map(r => r.funcArea))].sort(), []);
  const results = useMemo(() => {
    const q = query.trim().toLowerCase();
    return PD_ROLE_DATA.filter(r => {
      const aM = !filterArea || r.funcArea === filterArea;
      if (!q) return aM;
      return aM && (r.pdNum.toLowerCase().includes(q) || r.position.toLowerCase().includes(q) || r.grade.toLowerCase().includes(q) || r.funcArea.toLowerCase().includes(q) || r.roles.some(role => role.toLowerCase().includes(q)));
    });
  }, [query,filterArea]);

  const grouped = useMemo(() => {
    const g: Record<string, PDEntry[]> = {};
    results.forEach(r => { (g[r.funcArea] || (g[r.funcArea] = [])).push(r); });
    return g;
  }, [results]);

  const aC = (a: string) => FUNC_AREA_COLORS[a] || "#555";

  const Card = ({ r }: { r: PDEntry }) => (
    <div style={{border:"1px solid #e5e5e5",borderRadius:8,padding:"14px 16px",marginBottom:10,background:"white"}}>
      <div style={{display:"flex",flexWrap:"wrap",gap:8,alignItems:"flex-start",marginBottom:8}}>
        <span style={{fontSize:17,fontWeight:600,color:"#111",fontFamily:"monospace"}}>{r.pdNum}</span>
        <span style={{fontSize:13,fontWeight:500,color:"#333",alignSelf:"center"}}>{r.position}</span>
        <span style={{fontSize:12,color:"#666",alignSelf:"center",background:"#f0f0f0",padding:"2px 8px",borderRadius:20}}>{r.grade}</span>
        <span style={{fontSize:11,padding:"2px 9px",borderRadius:20,background:aC(r.funcArea)+"22",color:aC(r.funcArea),fontWeight:500,alignSelf:"center"}}>{r.funcArea}</span>
      </div>
      <div style={{fontSize:12,color:"#555",marginBottom:6,fontWeight:500}}>Required HR-Smart Roles:</div>
      <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
        {r.roles.map(role => <span key={role} style={{fontSize:12,padding:"3px 10px",borderRadius:20,background:"#E6F1FB",color:"#0C447C",border:"1px solid #c5d9f0"}}>{role}</span>)}
      </div>
    </div>
  );

  return (
    <div>
      <div style={{background:"#f0f6ff",border:"1px solid #c5d9f0",borderRadius:8,padding:"10px 14px",fontSize:12,marginBottom:16,color:"#185FA5"}}>
        <strong>Source:</strong> VHA WMC SOP Appendices B–O (November 2025). Search by PD#, position title, grade, functional area, or role name.
      </div>
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search PD#, position, grade, or role…" style={{flex:1,minWidth:200,padding:"7px 11px",fontSize:13,border:"1px solid #ddd",borderRadius:6}} autoFocus/>
        <select value={filterArea} onChange={e=>setFilterArea(e.target.value)} style={{padding:"7px 10px",fontSize:13,border:"1px solid #ddd",borderRadius:6}}>
          <option value="">All functional areas</option>
          {funcAreas.map(a => <option key={a} value={a}>{a}</option>)}
        </select>
        <div style={{display:"flex",border:"1px solid #ddd",borderRadius:6,overflow:"hidden"}}>
          {(["search","browse"] as const).map(m => (
            <button key={m} onClick={()=>setMode(m)} style={{padding:"6px 14px",fontSize:12,border:"none",cursor:"pointer",background:mode===m?"#185FA5":"white",color:mode===m?"white":"#555",textTransform:"capitalize"}}>{m}</button>
          ))}
        </div>
      </div>
      {results.length === 0 && <div style={{textAlign:"center",padding:"2.5rem",color:"#999",fontSize:13,border:"1px solid #e5e5e5",borderRadius:8}}>No PD records match.</div>}
      {mode === "search" && results.length > 0 && (
        <div>
          <div style={{fontSize:12,color:"#888",marginBottom:10}}>{results.length} record{results.length!==1?"s":""} found</div>
          {results.map((r,i) => <Card key={r.pdNum+r.funcArea+i} r={r}/>)}
        </div>
      )}
      {mode === "browse" && Object.keys(grouped).length > 0 && Object.entries(grouped).map(([area,rows]) => (
        <div key={area} style={{marginBottom:24}}>
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
            <div style={{width:4,height:20,borderRadius:2,background:aC(area)}}/>
            <span style={{fontSize:14,fontWeight:600,color:aC(area)}}>{area}</span>
            <span style={{fontSize:12,color:"#999"}}>({rows.length} PD{rows.length!==1?"s":""})</span>
          </div>
          {rows.map((r,i) => <Card key={r.pdNum+i} r={r}/>)}
        </div>
      ))}
    </div>
  );
}

// ── HR Role Compliance tab ────────────────────────────────────────────────────
interface EmployeeRow {
  [key: string]: unknown;
  "Empl ID"?: string;
  "Name"?: string;
  "User ID"?: string;
  "Email"?: string;
  "Station"?: string | number;
  "Location Desc"?: string;
  "POID"?: string | number;
  "PD/Functional Stmt Number"?: string;
  "Pay Plan"?: string;
  "Occ Series"?: string;
  "Grade"?: string;
  "Official Position Title"?: string;
  "Is Locked"?: number;
  "Last Signon Date"?: string | Date;
  "Role Name(s)"?: string;
}

interface ComplianceEntry extends PDEntry {
  missing: string[];
  extra: string[];
  compliant: boolean;
}

interface EmployeeCompliance {
  emplId: string;
  name: string;
  userId: string;
  email: string;
  station: string;
  loc: string;
  poid: string;
  pdNum: string;
  grade: string;
  title: string;
  isLocked: number;
  lastLogin: string | null;
  actualRoles: string[];
  byEntry: ComplianceEntry[];
  status: "compliant"|"missing"|"partial";
  funcAreas: string[];
}

function computeCompliance(emp: EmployeeRow, sopEntries: PDEntry[]): EmployeeCompliance {
  const actualRoles = ((emp["Role Name(s)"] || "") as string).split(",").map(r => r.trim()).filter(Boolean);
  const actualNormSet = new Set(actualRoles.map(normRole));

  const byEntry: ComplianceEntry[] = sopEntries.map(sop => {
    const missing = sop.roles.filter(r => !actualNormSet.has(normRole(r)));
    const expectedNormSet = new Set(sop.roles.map(normRole));
    const extra = actualRoles.filter(r => !isBase(r) && !expectedNormSet.has(normRole(r)));
    return { ...sop, missing, extra, compliant: missing.length === 0 };
  });

  const status: "compliant"|"missing"|"partial" = byEntry.some(e => e.compliant) ? "compliant"
    : byEntry.every(e => e.missing.length > 0) ? "missing" : "partial";

  const rawDate = emp["Last Signon Date"];
  const lastLogin = rawDate
    ? (rawDate instanceof Date ? rawDate : new Date(rawDate as string)).toISOString().split("T")[0]
    : null;

  return {
    emplId: String(emp["Empl ID"] || ""),
    name:   String(emp["Name"] || ""),
    userId: String(emp["User ID"] || ""),
    email:  String(emp["Email"] || ""),
    station: String(emp["Station"] || ""),
    loc:    String(emp["Location Desc"] || ""),
    poid:   String(emp["POID"] || ""),
    pdNum:  String(emp["PD/Functional Stmt Number"] || ""),
    grade:  `${emp["Pay Plan"] || ""}-${emp["Occ Series"] || ""}-${emp["Grade"] || ""}`,
    title:  String(emp["Official Position Title"] || ""),
    isLocked: Number(emp["Is Locked"] || 0),
    lastLogin,
    actualRoles,
    byEntry,
    status,
    funcAreas: [...new Set(sopEntries.map(e => e.funcArea))],
  };
}

function HRRoleCompliance() {
  const [loading, setLoading] = useState(false);
  const [loadErr, setLoadErr] = useState<string|null>(null);
  const [employees, setEmployees] = useState<EmployeeCompliance[]>([]);
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("");
  const [filterArea, setFilterArea] = useState("");
  const [filterStation, setFilterStation] = useState("");
  const [expanded, setExpanded] = useState<string|null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback((file: File) => {
    setLoading(true);
    setLoadErr(null);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buf = e.target?.result;
        const wb = XLSX.read(buf, { cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json<EmployeeRow>(ws, { defval: null });
        const processed = rows.map(emp => {
          const sopEntries = PD_ROLE_DATA.filter(s => s.pdNum === emp["PD/Functional Stmt Number"]);
          return computeCompliance(emp, sopEntries);
        });
        setEmployees(processed);
      } catch (err) {
        setLoadErr(err instanceof Error ? err.message : "Unknown error");
      } finally {
        setLoading(false);
      }
    };
    reader.onerror = () => { setLoadErr("Failed to read file"); setLoading(false); };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const handleInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
  }, [handleFile]);

  const allAreas = useMemo(() => [...new Set(employees.flatMap(e => e.funcAreas))].sort(), [employees]);
  const allStations = useMemo(() => [...new Set(employees.map(e => e.station))].sort(), [employees]);

  const filtered = useMemo(() => {
    const q = search.toLowerCase();
    return employees.filter(e =>
      (!filterStatus || e.status === filterStatus) &&
      (!filterArea || e.funcAreas.includes(filterArea)) &&
      (!filterStation || e.station === filterStation) &&
      (!q || [e.name,e.emplId,e.email,e.pdNum,e.title,e.userId].some(v => v && v.toLowerCase().includes(q)))
    );
  }, [employees,search,filterStatus,filterArea,filterStation]);

  const counts = useMemo(() => ({
    total: employees.length,
    compliant: employees.filter(e => e.status === "compliant").length,
    missing:   employees.filter(e => e.status === "missing").length,
    partial:   employees.filter(e => e.status === "partial").length,
  }), [employees]);

  const th: React.CSSProperties = {padding:"8px 10px",textAlign:"left",fontWeight:500,fontSize:12,color:"#666",borderBottom:"1px solid #e5e5e5",whiteSpace:"nowrap"};
  const td: React.CSSProperties = {padding:"8px 10px",verticalAlign:"top",fontSize:13};

  if (employees.length === 0) {
    return (
      <div>
        <div style={{background:"#f0f6ff",border:"1px solid #c5d9f0",borderRadius:8,padding:"10px 14px",fontSize:12,marginBottom:16,color:"#185FA5"}}>
          <strong>HR Role Compliance:</strong> Upload a Z_VAHR_ACCESS xlsx export to check each employee's HR-Smart roles against VHA WMC SOP requirements.
        </div>
        {loadErr && (
          <div style={{background:"#fff0f0",border:"1px solid #f0c0c0",borderRadius:8,padding:"1rem",color:"#A32D2D",fontSize:13,marginBottom:16}}>
            Failed to load file: {loadErr}
          </div>
        )}
        <div
          onDrop={handleDrop}
          onDragOver={e=>e.preventDefault()}
          onClick={()=>fileInputRef.current?.click()}
          style={{border:"2px dashed #c5d9f0",borderRadius:12,padding:"3rem 2rem",textAlign:"center",cursor:"pointer",background:"#f7faff",transition:"background 0.15s"}}
        >
          <div style={{fontSize:40,marginBottom:12}}>📂</div>
          <div style={{fontSize:16,fontWeight:600,color:"#185FA5",marginBottom:6}}>Drop your xlsx file here</div>
          <div style={{fontSize:13,color:"#888",marginBottom:16}}>or click to browse</div>
          <div style={{fontSize:12,color:"#aaa"}}>Expected format: Z_VAHR_ACCESS_VERSION42016.xlsx</div>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleInputChange} style={{display:"none"}}/>
        </div>
        {loading && <div style={{textAlign:"center",padding:"2rem",color:"#888",fontSize:13,marginTop:12}}>Processing file…</div>}
      </div>
    );
  }

  return (
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:"1rem"}}>
        <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,flex:1}}>
          {[
            {val:counts.total,     lbl:"Employees with SOP PD#",  color:"inherit"},
            {val:counts.compliant, lbl:"Fully compliant",          color:"#0C7A52"},
            {val:counts.partial,   lbl:"Partial match",            color:"#854F0B"},
            {val:counts.missing,   lbl:"Missing required roles",   color:"#A32D2D"},
          ].map(({val,lbl,color}) => (
            <div key={lbl} style={{background:"#f5f5f5",borderRadius:8,padding:"0.9rem",textAlign:"center"}}>
              <div style={{fontSize:26,fontWeight:500,color}}>{val}</div>
              <div style={{fontSize:11,color:"#666",marginTop:4}}>{lbl}</div>
            </div>
          ))}
        </div>
        <button onClick={()=>setEmployees([])} style={{marginLeft:12,padding:"6px 14px",fontSize:12,border:"1px solid #ddd",borderRadius:6,cursor:"pointer",background:"white",color:"#555",flexShrink:0}}>
          Change file
        </button>
      </div>

      <div style={{background:"#f0f6ff",border:"1px solid #c5d9f0",borderRadius:8,padding:"8px 14px",fontSize:12,marginBottom:14,color:"#185FA5"}}>
        <strong>How this works:</strong> Each employee's current HR-Smart roles are compared against the VHA WMC SOP required roles for their PD#.
        "Compliant" means all SOP-required roles are present. "Partial match" applies to PD#s shared across multiple appendices where at least one appendix is satisfied.
        Base roles (VA Employee, Benefits EBSS, etc.) are excluded from the "extra" analysis.
      </div>

      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name, Empl ID, PD#, email…"
          style={{flex:1,minWidth:180,padding:"6px 10px",fontSize:13,border:"1px solid #ddd",borderRadius:6}}/>
        <select value={filterStatus} onChange={e=>setFilterStatus(e.target.value)} style={{padding:"6px 10px",fontSize:13,border:"1px solid #ddd",borderRadius:6}}>
          <option value="">All statuses</option>
          <option value="compliant">✓ Compliant</option>
          <option value="partial">~ Partial match</option>
          <option value="missing">✗ Missing roles</option>
        </select>
        <select value={filterArea} onChange={e=>setFilterArea(e.target.value)} style={{padding:"6px 10px",fontSize:13,border:"1px solid #ddd",borderRadius:6}}>
          <option value="">All functional areas</option>
          {allAreas.map(a => <option key={a} value={a}>{a}</option>)}
        </select>
        <select value={filterStation} onChange={e=>setFilterStation(e.target.value)} style={{padding:"6px 10px",fontSize:13,border:"1px solid #ddd",borderRadius:6}}>
          <option value="">All stations</option>
          {allStations.map(s => <option key={s} value={s}>Stn {s}</option>)}
        </select>
      </div>

      <div style={{border:"1px solid #e5e5e5",borderRadius:10,overflow:"hidden"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
          <thead><tr style={{background:"#f9f9f9"}}>
            <th style={th}>Name / Email</th>
            <th style={th}>Empl ID</th>
            <th style={th}>PD#</th>
            <th style={th}>Title / Grade</th>
            <th style={th}>Functional Area</th>
            <th style={th}>Station</th>
            <th style={th}>Last Login</th>
            <th style={th}>Status</th>
          </tr></thead>
          <tbody>
            {filtered.length === 0
              ? <tr><td colSpan={8} style={{textAlign:"center",padding:"2rem",color:"#999"}}>No records match.</td></tr>
              : filtered.map(emp => {
                  const isOpen = expanded === emp.emplId;
                  const aC = (a: string) => FUNC_AREA_COLORS[a] || "#555";
                  return [
                    <tr key={emp.emplId}
                      onClick={()=>setExpanded(isOpen?null:emp.emplId)}
                      style={{borderBottom:isOpen?"none":"1px solid #f0f0f0",cursor:"pointer",
                        background:isOpen?"#f0f6ff":emp.status==="compliant"?"#f6fff9":emp.status==="missing"?"#fff8f6":"white"}}>
                      <td style={td}><div style={{fontWeight:500}}>{emp.name}</div><div style={{fontSize:11,color:"#888"}}>{emp.email}</div></td>
                      <td style={{...td,fontFamily:"monospace",fontSize:12}}>{emp.emplId}</td>
                      <td style={{...td,fontFamily:"monospace",fontWeight:600,fontSize:12}}>{emp.pdNum}</td>
                      <td style={td}><div style={{fontWeight:500,fontSize:12}}>{emp.title}</div><div style={{fontSize:11,color:"#888"}}>{emp.grade}</div></td>
                      <td style={td}>
                        {emp.funcAreas.map(a => (
                          <span key={a} style={{display:"inline-block",fontSize:11,padding:"1px 7px",borderRadius:20,background:aC(a)+"18",color:aC(a),marginRight:4,marginBottom:2,fontWeight:500}}>{a}</span>
                        ))}
                      </td>
                      <td style={td}><div style={{fontWeight:500}}>{emp.station}</div><div style={{fontSize:11,color:"#888"}}>{emp.loc}</div></td>
                      <td style={{...td,fontSize:12,color:emp.lastLogin?"#555":"#A32D2D"}}>
                        {emp.lastLogin || "Never"}
                        {emp.isLocked ? <div style={{fontSize:10,color:"#A32D2D"}}>locked</div> : null}
                      </td>
                      <td style={td}><StatusPill status={emp.status}/></td>
                    </tr>,
                    isOpen && (
                      <tr key={emp.emplId+"_detail"} style={{borderBottom:"1px solid #e5e5e5"}}>
                        <td colSpan={8} style={{padding:"0 0 16px 0",background:"#f7faff"}}>
                          <div style={{padding:"12px 18px"}}>
                            <div style={{marginBottom:14}}>
                              <div style={{fontSize:12,fontWeight:600,color:"#333",marginBottom:6}}>
                                Current HR-Smart roles ({emp.actualRoles.length}):
                              </div>
                              <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                                {emp.actualRoles.map(r => (
                                  <span key={r} style={{fontSize:11,padding:"2px 8px",borderRadius:20,background:"#f0f0f0",color:"#444",border:"1px solid #ddd"}}>{r}</span>
                                ))}
                              </div>
                            </div>
                            {emp.byEntry.map((entry,i) => (
                              <div key={i} style={{marginBottom:14,border:"1px solid #e0eaff",borderRadius:8,overflow:"hidden"}}>
                                <div style={{padding:"8px 14px",background:FUNC_AREA_COLORS[entry.funcArea]||"#555",color:"white",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                                  <span style={{fontWeight:600,fontSize:13}}>{entry.funcArea} — {entry.position} ({entry.grade})</span>
                                  <StatusPill status={entry.compliant?"compliant":"missing"}/>
                                </div>
                                <div style={{padding:"12px 14px",background:"white"}}>
                                  <div style={{fontSize:12,fontWeight:600,color:"#333",marginBottom:6}}>Required roles:</div>
                                  <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:entry.missing.length||entry.extra.length?10:0}}>
                                    {entry.roles.map(role => {
                                      const has = new Set(emp.actualRoles.map(normRole)).has(normRole(role));
                                      return (
                                        <span key={role} style={{fontSize:12,padding:"3px 10px",borderRadius:20,
                                          background:has?"#e6f7ee":"#fff0e6",
                                          color:has?"#0C7A52":"#A32D2D",
                                          border:`1px solid ${has?"#a3d9b8":"#f0c0a0"}`,
                                          display:"flex",alignItems:"center",gap:4}}>
                                          <span>{has?"✓":"✗"}</span>{role}
                                        </span>
                                      );
                                    })}
                                  </div>
                                  {entry.missing.length > 0 && (
                                    <div style={{fontSize:12,color:"#A32D2D",marginTop:4}}>
                                      <strong>Missing ({entry.missing.length}):</strong> {entry.missing.join(", ")}
                                    </div>
                                  )}
                                  {entry.extra.length > 0 && (
                                    <div style={{fontSize:12,color:"#854F0B",marginTop:4}}>
                                      <strong>Extra / unexpected ({entry.extra.length}):</strong> {entry.extra.join(", ")}
                                    </div>
                                  )}
                                </div>
                              </div>
                            ))}
                          </div>
                        </td>
                      </tr>
                    )
                  ];
                })
            }
          </tbody>
        </table>
      </div>
      <p style={{fontSize:12,color:"#999",marginTop:8,textAlign:"right"}}>
        Showing {filtered.length} of {employees.length} employees with SOP PD#
      </p>
    </div>
  );
}

// ── Root App ─────────────────────────────────────────────────────────────────
export default function VAStaleAccess() {
  const [tab, setTab] = useState(0);
  const [exceptions, setExceptions] = useState<Record<string, ExceptionRecord>>({});
  const [saving, setSaving] = useState(false);
  const [storageOk, setStorageOk] = useState(true);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) setExceptions(JSON.parse(raw));
    } catch (_) {
      setStorageOk(false);
    }
  }, []);

  const persist = useCallback((updated: Record<string, ExceptionRecord>) => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
    } catch (_) {
      setStorageOk(false);
    }
  }, []);

  const handleSave = useCallback((emplId: string, form: ExceptionFormValues) => {
    setSaving(true);
    const updated = {...exceptions, [emplId]: {...form, savedAt: TODAY.toISOString().split("T")[0]}};
    setExceptions(updated);
    persist(updated);
    setSaving(false);
  }, [exceptions, persist]);

  const handleRemove = useCallback((emplId: string) => {
    const updated = {...exceptions};
    delete updated[emplId];
    setExceptions(updated);
    persist(updated);
  }, [exceptions, persist]);

  const staleExcepted = STALE_DATA.filter(d => isActive(exceptions[d.emplId])).length;
  const mismatchExcepted = MISMATCH_DATA.filter(d => isActive(exceptions[d.emplId])).length;
  const TABS = ["Stale logins","Scope mismatch","HR Role Lookup","HR Role Compliance"];

  return (
    <div style={{padding:"1rem 0",fontFamily:"sans-serif"}}>
      {!storageOk && (
        <div style={{background:"#FFF3CD",border:"1px solid #FFD700",borderRadius:8,padding:"8px 12px",fontSize:12,marginBottom:12,color:"#856404"}}>
          Storage unavailable — exceptions will not persist.
        </div>
      )}
      <div style={{display:"flex",gap:4,marginBottom:"1.5rem",borderBottom:"1px solid #e5e5e5"}}>
        {TABS.map((t,i) => (
          <button key={t} onClick={()=>setTab(i)} style={{padding:"8px 18px",fontSize:13,border:"none",borderBottom:tab===i?"2px solid #185FA5":"2px solid transparent",borderRadius:0,cursor:"pointer",background:"transparent",color:tab===i?"#185FA5":"#666",fontWeight:tab===i?500:400}}>{t}</button>
        ))}
      </div>

      {tab === 0 && (
        <>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:"1.5rem"}}>
            {[
              {val:TOTAL_ADMIN,                         lbl:"Total VA Admin Officer users",color:"inherit"},
              {val:STALE_DATA.length,                   lbl:"No login in 6+ months",       color:"#854F0B"},
              {val:staleExcepted,                       lbl:"Active exceptions",            color:"#185FA5"},
              {val:STALE_DATA.length-staleExcepted,     lbl:"Need review",                 color:"#A32D2D"},
            ].map(({val,lbl,color}) => (
              <div key={lbl} style={{background:"#f5f5f5",borderRadius:8,padding:"0.9rem",textAlign:"center"}}>
                <div style={{fontSize:26,fontWeight:500,color}}>{val}</div>
                <div style={{fontSize:11,color:"#666",marginTop:4}}>{lbl}</div>
              </div>
            ))}
          </div>
          <DataTable data={STALE_DATA} exceptions={exceptions} onSave={handleSave} onRemove={handleRemove} saving={saving} showMismatchCount={false}/>
        </>
      )}

      {tab === 1 && (
        <>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:"1.5rem"}}>
            {[
              {val:TOTAL_ADMIN,                           lbl:"Total VA Admin Officer users",color:"inherit"},
              {val:MISMATCH_DATA.length,                  lbl:"Users with scope mismatch",   color:"#854F0B"},
              {val:mismatchExcepted,                      lbl:"Active exceptions",            color:"#185FA5"},
              {val:MISMATCH_DATA.length-mismatchExcepted, lbl:"Need review",                 color:"#A32D2D"},
            ].map(({val,lbl,color}) => (
              <div key={lbl} style={{background:"#f5f5f5",borderRadius:8,padding:"0.9rem",textAlign:"center"}}>
                <div style={{fontSize:26,fontWeight:500,color}}>{val}</div>
                <div style={{fontSize:11,color:"#666",marginTop:4}}>{lbl}</div>
              </div>
            ))}
          </div>
          <div style={{background:"#f0f6ff",border:"1px solid #c5d9f0",borderRadius:8,padding:"8px 14px",fontSize:12,marginBottom:12,color:"#185FA5"}}>
            <strong>Mismatch logic:</strong> Org Cd roles flagged if code ≠ user's Org Code. POI roles flagged if POI ≠ user's POID.
          </div>
          <DataTable data={MISMATCH_DATA} exceptions={exceptions} onSave={handleSave} onRemove={handleRemove} saving={saving} showMismatchCount={true}/>
        </>
      )}

      {tab === 2 && <HRRoleLookup/>}
      {tab === 3 && <HRRoleCompliance/>}
    </div>
  );
}
