package domain;
import java.util.ArrayList;

public class NemoRecordRow {
	public String Time;
	public double Lat;
	public double Lon;
	public String CellID;
//SEARFCN	SPCI	SRSCP
	public int SEARFCN;
	public int SPCI;
	public double SRSCP;
	
//	NEARFCN	NPCI	NRSCP 1 to 8
	public ArrayList<Integer> NEARFCNList;
	public ArrayList<Integer> NPCIList;
	public ArrayList<Double> NRSCPList;
	public String getTime() {
		return Time;
	}
	public double getLat() {
		return Lat;
	}
	public double getLon() {
		return Lon;
	}
	public String getCellID() {
		return CellID;
	}
	public int getSEARFCN() {
		return SEARFCN;
	}
	public int getSPCI() {
		return SPCI;
	}
	public double getSRSCP() {
		return SRSCP;
	}
	public ArrayList<Integer> getNEARFCNList() {
		return NEARFCNList;
	}
	public ArrayList<Integer> getNPCIList() {
		return NPCIList;
	}
	public ArrayList<Double> getNRSCPList() {
		return NRSCPList;
	}
	public void setTime(String time) {
		Time = time;
	}
	public void setLat(double lat) {
		Lat = lat;
	}
	public void setLon(double lon) {
		Lon = lon;
	}
	public void setCellID(String cellID) {
		CellID = cellID;
	}
	public void setSEARFCN(int sEARFCN) {
		SEARFCN = sEARFCN;
	}
	public void setSPCI(int sPCI) {
		SPCI = sPCI;
	}
	public void setSRSCP(double sRSCP) {
		SRSCP = sRSCP;
	}
	public void setNEARFCNList(ArrayList<Integer> nEARFCNList) {
		NEARFCNList = nEARFCNList;
	}
	public void setNPCIList(ArrayList<Integer> nPCIList) {
		NPCIList = nPCIList;
	}
	public void setNRSCPList(ArrayList<Double> nRSCPList) {
		NRSCPList = nRSCPList;
	}
	
	
}