package net.tugdev.java2excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Queue;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class ExcelTabbedObjectPrinter {

	private Workbook workbook;
	private CellStyle headerCellStyle;
	private CellStyle idCellStyle;
	private CellStyle dateCellStyle;
	private Map<Sheet, Integer> currentRows;
	private Map<String, Integer> objectsCount;
	private Map<String, String> sheetNames;
	private Map<Integer, String> processedObjects;
	private Set<Integer> printedObjects;
	private Map<Class<?>, List<Field>> fieldsCache;
	private Map<Sheet, Queue<Object>> objectsQueue;

	private static Comparator<Field> fieldComparator = new Comparator<Field>() {
		@Override
		public int compare(Field field1, Field field2) {
			return field1.getName().compareTo(field2.getName());
		}
	};

	public ExcelTabbedObjectPrinter() {
		this.workbook = new SXSSFWorkbook(1000);

		this.headerCellStyle = workbook.createCellStyle();
		this.headerCellStyle.setFillForegroundColor(HSSFColor.YELLOW.index);
		this.headerCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

		this.idCellStyle = workbook.createCellStyle();
		this.idCellStyle
				.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
		this.idCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

		this.dateCellStyle = workbook.createCellStyle();
		CreationHelper createHelper = workbook.getCreationHelper();
		this.dateCellStyle.setDataFormat(createHelper.createDataFormat()
				.getFormat("m/d/yy"));

		this.currentRows = new HashMap<Sheet, Integer>();
		this.objectsCount = new HashMap<String, Integer>();
		this.sheetNames = new HashMap<String, String>();
		this.processedObjects = new HashMap<Integer, String>();
		this.fieldsCache = new HashMap<Class<?>, List<Field>>();
		this.objectsQueue = new HashMap<Sheet, Queue<Object>>();
		this.printedObjects = new HashSet<Integer>();
	}

	private Sheet getSheetForClass(Class<?> classInstance) {
		String fqn = classInstance.getName();
		String sheetName = this.sheetNames.get(fqn);
		if (sheetName == null) {
			String className;
			if(classInstance.isAnonymousClass()) {
				className = classInstance.getSuperclass().getSimpleName();
			} else {
				className = classInstance.getSimpleName();
			}
			className = unCamelize(className);
			className.substring(0, Math.min(className.length(), 28));
			if (workbook.getSheet(className) == null) {
				sheetName = className;
			} else {
				int inc = 1;
				do {
					sheetName = className + "~" + inc;
					inc++;
				} while (workbook.getSheet(sheetName) != null);
			}
			Sheet sheet = workbook.createSheet(sheetName);
			this.sheetNames.put(fqn, sheetName);
			this.currentRows.put(sheet, 0);
			this.objectsCount.put(fqn, 1);
			this.objectsQueue.put(sheet, new LinkedList<Object>());
			return sheet;
		}
		return workbook.getSheet(sheetName);
	}

	public static String unCamelize(String inputString) {
		Pattern p = Pattern.compile("\\p{Lu}");
		Matcher m = p.matcher(inputString);
		StringBuffer sb = new StringBuffer();
		while (m.find()) {
			m.appendReplacement(sb, " " + m.group());
		}
		m.appendTail(sb);
		return sb.toString().trim();
	}

	public String addObject(Object object) {
		if(object instanceof Collection) {
			return addCollection((Collection<?>)object);
		}
		if(object instanceof Map) {
			return addMap((Map<?,?>)object);
		}
		Sheet worksheet = getSheetForClass(object.getClass());
		Integer id = System.identityHashCode(object);
		String objectId = this.processedObjects.get(id);
		if(objectId == null) {
			objectId = cacheObjectId(object);
			this.processedObjects.put(id, objectId);
			addObject(worksheet, objectId, object);
			this.printedObjects.add(id);
			processQueue(worksheet);
		} else if(!this.printedObjects.contains(id)) {
			this.objectsQueue.get(worksheet).offer(object);
		}
		return objectId;
	}
	
	public String addCollection(Collection<?> collection) {
		for(Object object : collection) {
			addObject(object);
		}
		return null;
	}
	
	public String addMap(Map<?,?> map) {
		MapWrapper wrapper = new MapWrapper();
		wrapper.map = map;
		return addObject(wrapper);
	}

	private void addObject(Sheet worksheet, String objectId, Object object) {
		Class<?> objectClass = object.getClass();
		if (!hasHeader(worksheet)) {
			addHeader(worksheet, objectClass);
		}
		int x = 0;
		int y = currentRow(worksheet);
		Cell cell = setValue(worksheet, x, y, objectId);
		cell.setCellStyle(idCellStyle);
		x++;
		List<Field> fields = getAllFields(objectClass);
		for (Field field : fields) {
			field.setAccessible(true);
			try {
				Object value = field.get(object);
				setValue(worksheet, x, y, value);
			} catch (IllegalArgumentException e) {
				e.printStackTrace();
			} catch (IllegalAccessException e) {
				e.printStackTrace();
			}
			x++;
		}
		nextRow(worksheet);
	}

	private void addHeader(Sheet worksheet, Class<?> objecType) {
		List<Field> fields = getAllFields(objecType);
		int x = 0;
		int y = currentRow(worksheet);
		Cell idCell = getCell(worksheet, x, y);
		idCell.setCellValue("ID " + this.sheetNames.get(objecType.getName()));
		idCell.setCellStyle(idCellStyle);
		x++;
		for (Field field : fields) {
			if(Map.class.isAssignableFrom(field.getType())) {
				Cell cell = setValue(worksheet, x, y, field.getName()+" keys");
				cell.setCellStyle(headerCellStyle);
				x++;
				cell = setValue(worksheet, x, y, field.getName()+" values");
				cell.setCellStyle(headerCellStyle);
				x++;
			} else {
				Cell cell = setValue(worksheet, x, y, field.getName());
				cell.setCellStyle(headerCellStyle);
				x++;
			}
		}
		y = nextRow(worksheet);
		x = 1;
		for (Field field : fields) {
			if(Map.class.isAssignableFrom(field.getType())) {
				x++;
				x++;
			} else {
				Cell cell = setValue(worksheet, x, y, field.getType().getName());
				cell.setCellStyle(headerCellStyle);
				x++;
			}
		}
		nextRow(worksheet);
	}

	private boolean hasHeader(Sheet worksheet) {
		return this.currentRows.get(worksheet) > 0;
	}

	private int nextRow(Sheet worksheet) {
		int nextRow = this.currentRows.get(worksheet) + 1;
		this.currentRows.put(worksheet, nextRow);
		return nextRow;
	}

	private int currentRow(Sheet worksheet) {
		return this.currentRows.get(worksheet);
	}

	private int endRow(Sheet worksheet) {
		return this.currentRows.put(worksheet, worksheet.getLastRowNum() + 1);
	}

	private Cell getCell(Sheet worksheet, int x, int y) {
		Row row = worksheet.getRow(y);
		if (row == null) {
			row = worksheet.createRow(y);
		}
		Cell cell = row.getCell(x);
		if (cell == null) {
			cell = row.createCell(x);
		}
		return cell;
	}

	private String cacheObjectId(Object object) {
		String fqn = object.getClass().getName();
		String className = this.sheetNames.get(fqn);
		int count = objectsCount.get(fqn);
		objectsCount.put(fqn, count + 1);
		return className + " " + count;
	}

	private Cell setValue(Sheet worksheet, int x, int y, Object value) {
		Cell cell = getCell(worksheet, x, y);
		if (value == null)
			return cell;
		if (value instanceof Number) {
			cell.setCellValue(((Number) value).doubleValue());
		} else if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
			cell.setCellStyle(dateCellStyle);
		} else if (value instanceof Calendar) {
			cell.setCellValue((Calendar) value);
			cell.setCellStyle(dateCellStyle);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Collection) {
			for (Object element : (Collection<?>) value) {
				setValue(worksheet, x, y, element);
				y++;
			}
			endRow(worksheet);
		} else if(value.getClass().isArray()) {
			for (int i=0, len = java.lang.reflect.Array.getLength(value); i<len; i++) {
				Object obj = java.lang.reflect.Array.get(value, i);
				setValue(worksheet, x, y, obj);
				y++;
			}
			endRow(worksheet);
		} else if (value instanceof Map) {
			for (Map.Entry<?, ?> entry : ((Map<?, ?>) value).entrySet()) {
				setValue(worksheet, x, y, entry.getKey());
				setValue(worksheet, x + 1, y, entry.getValue());
				y++;
			}
			endRow(worksheet);
		} else {
			String objectId = addObject(value);
			cell.setCellValue(objectId);
			cell.setCellStyle(idCellStyle);
		}
		return cell;
	}
	
	private void processQueue(Sheet worksheet) {
		Queue<Object> q = objectsQueue.get(worksheet);
		while(!q.isEmpty()) {
			addObject(q.poll());
		}
	}

	private List<Field> getAllFields(Class<?> type) {
		List<Field> fields = this.fieldsCache.get(type);
		if(fields != null) return fields;
		fields = new ArrayList<Field>();
		for (Class<?> c = type; c != null; c = c.getSuperclass()) {
			List<Field> df = Arrays.asList(c.getDeclaredFields());
			Iterator<Field> fit = df.iterator();
			while(fit.hasNext()) {
				Field field = fit.next();
				if(!java.lang.reflect.Modifier.isStatic(field.getModifiers())) {
					fields.add(field);
				}
			}
		}
		Collections.sort(fields, fieldComparator);
		this.fieldsCache.put(type, fields);
		return fields;
	}
	
	public void autoSizeAll() {
		int nbSheets = workbook.getNumberOfSheets();
		for(int i=0; i<nbSheets; i++) {
			autoSizeSheet(this.workbook.getSheetAt(i));
		}
	}
	
	private void autoSizeSheet(Sheet worksheet) {
		Row row = worksheet.getRow(0);
		if(row != null) {
			int maxCol = row.getLastCellNum();
			for(int i=0; i<=maxCol; i++) {
				worksheet.autoSizeColumn(i);
			}
		}
	}

	public void save(OutputStream fileOut) throws IOException {
		try {
			workbook.write(fileOut);
			fileOut.flush();
			fileOut.close();
		} catch (IOException e) {
			throw e;
		} finally {
			try {
				if (fileOut != null) {
					fileOut.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	public void save(String filePath) throws IOException {
		save(new FileOutputStream(filePath));
	}
	
	public void save(File file) throws IOException {
		save(new FileOutputStream(file));
	}

}
