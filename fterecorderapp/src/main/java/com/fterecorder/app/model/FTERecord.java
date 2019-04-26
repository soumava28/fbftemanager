package com.fterecorder.app.model;

import java.time.LocalDate;
import java.util.Date;

import org.springframework.data.annotation.Id;
import org.springframework.data.annotation.PersistenceConstructor;
import org.springframework.data.annotation.Transient;
import org.springframework.data.mongodb.core.mapping.Document;
import org.springframework.format.annotation.DateTimeFormat;
import org.springframework.format.annotation.DateTimeFormat.ISO;

@Document(collection = "FTERecord")
public class FTERecord {

	@Transient
	public static final String SEQUENCE_NAME = "fterecord_seq";
	@Id
	long id;
	String track;
	int ftesPerRosterCount;
	int defectCount;
	int widgetCount;
	int ftesLoanedCount;
	int ptoCount;
	int ftesBorrowedCount;
	int ftesForPerformanceCount;
	int ftesForExtendedScenarioExecCount;
	int excessFteCount;
	int medsDefAnalysisCount;
	int workableDefectsCount;
	int defectFteCount;

	public int getDefectFteCount() {
		return defectFteCount;
	}

	public void setDefectFteCount(int defectFteCount) {
		this.defectFteCount = defectFteCount;
	}

	public int getWorkableDefectsCount() {
		return workableDefectsCount;
	}

	public void setWorkableDefectsCount(int workableDefectsCount) {
		this.workableDefectsCount = workableDefectsCount;
	}

	@DateTimeFormat(iso = DateTimeFormat.ISO.DATE)
	LocalDate createDt;
	@DateTimeFormat(iso = DateTimeFormat.ISO.DATE)
	LocalDate weekStDt;
	@DateTimeFormat(iso = DateTimeFormat.ISO.DATE)
	LocalDate weekEdDt;
	String note;
	Employee employee;

	public FTERecord() {
	}

	@PersistenceConstructor
	public FTERecord(long id, String track, int ftesPerRosterCount, int defectCount, int defectFteCount,
			int widgetCount, int ftesLoanedCount, int ptoCount, int ftesBorrowedCount, int ftesForPerformanceCount,
			int ftesForExtendedScenarioExecCount, int excessFteCount, int workableDefectsCount,
			int medsDefAnalysisCount, @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) LocalDate createDt, @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) LocalDate weekStDt, @DateTimeFormat(iso = DateTimeFormat.ISO.DATE) LocalDate weekEdDt, String note) {

		super();
		this.id = id;
		this.track = track;
		this.ftesPerRosterCount = ftesPerRosterCount;
		this.defectCount = defectCount;
		this.defectFteCount = defectFteCount;
		this.widgetCount = widgetCount;
		this.ftesLoanedCount = ftesLoanedCount;
		this.ptoCount = ptoCount;
		this.ftesBorrowedCount = ftesBorrowedCount;
		this.ftesForPerformanceCount = ftesForPerformanceCount;
		this.ftesForExtendedScenarioExecCount = ftesForExtendedScenarioExecCount;
		this.excessFteCount = excessFteCount;
		this.workableDefectsCount = workableDefectsCount;
		this.medsDefAnalysisCount = medsDefAnalysisCount;
		this.createDt = createDt;
		this.weekStDt = weekStDt;
		this.weekEdDt = weekEdDt;
		this.note = note;
	}

	public long getId() {
		return id;
	}

	public void setId(long id) {
		this.id = id;
	}

	public String getTrack() {
		return track;
	}

	public void setTrack(String track) {
		this.track = track;
	}

	public int getFtesPerRosterCount() {
		return ftesPerRosterCount;
	}

	public void setFtesPerRosterCount(int ftesPerRosterCount) {
		this.ftesPerRosterCount = ftesPerRosterCount;
	}

	public int getDefectCount() {
		return defectCount;
	}

	public void setDefectCount(int defectCount) {
		this.defectCount = defectCount;
	}

	public int getWidgetCount() {
		return widgetCount;
	}

	public void setWidgetCount(int widgetCount) {
		this.widgetCount = widgetCount;
	}

	public int getFtesLoanedCount() {
		return ftesLoanedCount;
	}

	public void setFtesLoanedCount(int ftesLoanedCount) {
		this.ftesLoanedCount = ftesLoanedCount;
	}

	public int getPtoCount() {
		return ptoCount;
	}

	public void setPtoCount(int ptoCount) {
		this.ptoCount = ptoCount;
	}

	public int getFtesBorrowedCount() {
		return ftesBorrowedCount;
	}

	public void setFtesBorrowedCount(int ftesBorrowedCount) {
		this.ftesBorrowedCount = ftesBorrowedCount;
	}

	public int getFtesForPerformanceCount() {
		return ftesForPerformanceCount;
	}

	public void setFtesForPerformanceCount(int ftesForPerformanceCount) {
		this.ftesForPerformanceCount = ftesForPerformanceCount;
	}

	public int getFtesForExtendedScenarioExecCount() {
		return ftesForExtendedScenarioExecCount;
	}

	public void setFtesForExtendedScenarioExecCount(int ftesForExtendedScenarioExecCount) {
		this.ftesForExtendedScenarioExecCount = ftesForExtendedScenarioExecCount;
	}

	public int getExcessFteCount() {
		return excessFteCount;
	}

	public void setExcessFteCount(int excessFteCount) {
		this.excessFteCount = excessFteCount;
	}

	public int getMedsDefAnalysisCount() {
		return medsDefAnalysisCount;
	}

	public void setMedsDefAnalysisCount(int medsDefAnalysisCount) {
		this.medsDefAnalysisCount = medsDefAnalysisCount;
	}

	public LocalDate getCreateDt() {
		return createDt;
	}

	public void setCreateDt(LocalDate createDt) {
		this.createDt = createDt;
	}

	public LocalDate getWeekStDt() {
		return weekStDt;
	}

	public void setWeekStDt(LocalDate weekStDt) {
		this.weekStDt = weekStDt;
	}

	public LocalDate getWeekEdDt() {
		return weekEdDt;
	}

	public void setWeekEdDt(LocalDate weekEdDt) {
		this.weekEdDt = weekEdDt;
	}
	
	public String getNote() {
		return note;
	}

	public void setNote(String note) {
		this.note = note;
	}

	public Employee getEmployee() {
		return employee;
	}

	public void setEmployee(Employee employee) {
		this.employee = employee;
	}

}
