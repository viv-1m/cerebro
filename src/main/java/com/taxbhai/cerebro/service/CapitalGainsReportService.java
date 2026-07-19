package com.taxbhai.cerebro.service;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

import com.taxbhai.cerebro.constants.CapitalGainsConstants;

/**
 * Parses Zerodha ("taxpnl-*.xlsx") and Groww ("Stocks_Capital_Gains_Report_*.xlsx")
 * capital-gains statements and produces one consolidated P&L row per
 * (security, security type, trade term, quarter, broker).
 *
 * NOTE ON FORMAT DETECTION: this is built against the two sample files you
 * uploaded. Both brokers occasionally tweak their export layout, so treat the
 * section-header strings below (SHORT_TERM_HEADERS etc.) as the first thing
 * to check if a future export stops parsing correctly.
 *
 * NOTE ON SECURITY TYPE: neither broker's CG statement carries an explicit
 * "Equity / Mutual Fund / Other" column outside of Zerodha's sheet layout, so
 * for Groww (and for Zerodha's "Non-equity" bucket) the type is inferred from
 * the instrument name via {@link #classifyByName(String)}. Adjust the
 * keyword lists there if you spot a misclassified row.
 *
 * NOTE ON HOLDING-PERIOD THRESHOLDS: see {@link CapitalGainsConstants} —
 * gold no longer has a special-cased threshold; every instrument type uses
 * the same {@code LONG_TERM_HOLDING_DAYS_THRESHOLD}.
 */
@Service
public class CapitalGainsReportService {
    public CapitalGainsReportService() {
        super();
    }

    @Override
    public int hashCode() {
        return super.hashCode();
    }

    // ---------- Public output model ----------

    public record CapitalGainRow(
            String securityName,
            String securityType,   // Equity | Mutual Fund | Gold
            String tradeTerm,      // Short Term | Long Term
            BigDecimal profitLossAmount,
            String dematProvider,  // zerodha | groww
            BigDecimal buyValue,
            BigDecimal sellValue,
            String quarter         // Q1 | Q2 | Q3 | Q4
    ) {}

    private enum Provider { ZERODHA, GROWW, UNKNOWN }

    // Internal wrapper to track quarter during processing
    private record CapitalGainRowWithQuarter(CapitalGainRow row, String quarter) {}

    // Holds consolidated values for a single security across multiple trades
    private record ConsolidatedRowData(
            BigDecimal profitLossAmount,
            BigDecimal buyValue,
            BigDecimal sellValue
    ) {}

    // Intermediate record for parsed row data before consolidation
    private record ParsedRow(
            String securityName,
            String securityType,
            String tradeTerm,
            BigDecimal profitLossAmount,
            String dematProvider,
            BigDecimal buyValue,
            BigDecimal sellValue,
            String quarter
    ) {}

    // Grouping key used to consolidate duplicate rows into one.
    private record Key(String provider, String securityName, String securityType,
                       String tradeTerm, String quarter) {}

    // ---------- Entry point ----------

    /**
     * Accepts one or more capital-gains xlsx statements (Zerodha and/or
     * Groww, in any mix/order) and returns the consolidated P&L rows grouped
     * by trade category (Short Term, Long Term, Intraday, F&O).
     * Returns a map: {@code {"shortTerm": [...], "longTerm": [...], ...}}.
     * Each row includes a quarter field (Q1, Q2, Q3, Q4).
     */
    public Map<String, List<CapitalGainRow>> generateConsolidatedReport(List<MultipartFile> files) throws IOException {
        Map<Key, ConsolidatedRowData> consolidated = new LinkedHashMap<>();

        for (MultipartFile file : files) {
            if (file == null || file.isEmpty()) {
                continue;
            }
            try (InputStream is = file.getInputStream();
                 Workbook workbook = WorkbookFactory.create(is)) {

                Provider provider = detectProvider(workbook);
                List<ParsedRow> rows = switch (provider) {
                    case ZERODHA -> parseZerodha(workbook);
                    case GROWW -> parseGroww(workbook);
                    case UNKNOWN -> throw new IllegalArgumentException(
                            "Unrecognized report format for file: " + file.getOriginalFilename());
                };

                for (ParsedRow row : rows) {
                    Key key = new Key(row.dematProvider(), row.securityName(),
                            row.securityType(), row.tradeTerm(), row.quarter());
                    ConsolidatedRowData current = consolidated.get(key);
                    if (current == null) {
                        consolidated.put(key, new ConsolidatedRowData(
                                row.profitLossAmount(),
                                row.buyValue(),
                                row.sellValue()));
                    } else {
                        consolidated.put(key, new ConsolidatedRowData(
                                current.profitLossAmount().add(row.profitLossAmount()),
                                current.buyValue().add(row.buyValue()),
                                current.sellValue().add(row.sellValue())));
                    }
                }
            }
        }

        // Initialize flat map structure: category -> list of rows
        Map<String, List<CapitalGainRow>> result = new LinkedHashMap<>();
        result.put("shortTerm", new ArrayList<>());
        result.put("longTerm", new ArrayList<>());
        result.put("intraday", new ArrayList<>());
        result.put("fno", new ArrayList<>());

        // Populate the flat structure, sorted by provider and security name
        consolidated.entrySet().stream()
                .map(e -> new CapitalGainRow(
                        truncateSecurityName(e.getKey().securityName()),
                        e.getKey().securityType(),
                        e.getKey().tradeTerm(),
                        e.getValue().profitLossAmount().setScale(2, RoundingMode.HALF_UP),
                        e.getKey().provider(),
                        e.getValue().buyValue().setScale(2, RoundingMode.HALF_UP),
                        e.getValue().sellValue().setScale(2, RoundingMode.HALF_UP),
                        e.getKey().quarter()))
                .sorted(Comparator.comparing(CapitalGainRow::dematProvider)
                        .thenComparing(CapitalGainRow::securityName))
                .forEach(row -> {
                    String category = getCategoryKey(row.tradeTerm());
                    if (result.containsKey(category)) {
                        result.get(category).add(row);
                    }
                });

        return result;
    }


    private String getCategoryKey(String tradeTerm) {
        return switch (tradeTerm) {
            case "Short Term" -> "shortTerm";
            case "Long Term" -> "longTerm";
            case "Intraday" -> "intraday";
            case "F&O" -> "fno";
            default -> "other";
        };
    }

    // ---------- Provider detection ----------

    private Provider detectProvider(Workbook workbook) {
        for (Sheet sheet : workbook) {
            if (sheet.getSheetName().startsWith("Tradewise Exits")) {
                return Provider.ZERODHA;
            }
        }
        Sheet sheet1 = workbook.getSheet("Sheet1");
        if (sheet1 != null) {
            for (int r = 0; r <= Math.min(10, sheet1.getLastRowNum()); r++) {
                String cell0 = getString(sheet1, r, 0);
                if (cell0 != null && cell0.contains("Capital Gains Statement")) {
                    return Provider.GROWW;
                }
            }
        }
        return Provider.UNKNOWN;
    }

    // ---------- Zerodha parser ----------

    private static final Map<String, String> ZERODHA_SECTION_TERM = Map.of(
            "Equity - Short Term", "Short Term",
            "Equity - Long Term", "Long Term"
    );
    // Sections we deliberately skip: Intraday (business income) and Buyback
    // (taxed as deemed dividend, not capital gains under current Indian tax rules).
    private static final Set<String> ZERODHA_SKIP_SECTIONS = Set.of(
            "Equity - Intraday", "Equity - Buyback", "F&O", "Currency", "Commodity"
    );

    private List<ParsedRow> parseZerodha(Workbook workbook) {
        List<ParsedRow> results = new ArrayList<>();
        Sheet sheet = workbook.getSheetIndex("Tradewise Exits from 2025-04-01") >= 0
                ? workbook.getSheet("Tradewise Exits from 2025-04-01")
                : findSheetStartingWith(workbook, "Tradewise Exits");
        if (sheet == null) return results;

        String currentSection = null;
        List<String> header = null;

        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            String col1 = getString(sheet, r, 1);
            if (col1 == null || col1.isBlank()) {
                continue;
            }
            if (isZerodhaSectionHeader(col1)) {
                currentSection = col1;
                header = null;
                continue;
            }
            if ("Symbol".equals(col1)) {
                header = readRowAsStrings(sheet, r);
                continue;
            }
            if (currentSection == null || header == null) {
                continue; // stray text (title/disclaimer rows)
            }
            if (ZERODHA_SKIP_SECTIONS.contains(currentSection)) {
                continue;
            }

            int exitDateIdx = header.indexOf("Exit Date");
            int profitIdx = header.indexOf("Profit");
            int holdingIdx = header.indexOf("Period of Holding");
            int buyValueIdx = header.indexOf("Buy Value");
            int sellValueIdx = header.indexOf("Sell Value");
            if (exitDateIdx < 0 || profitIdx < 0) continue;

            String symbol = col1; // Symbol is in column index 1 for this sheet
            LocalDate exitDate = parseDate(sheet, r, exitDateIdx);
            BigDecimal profit = getBigDecimal(sheet, r, profitIdx);
            BigDecimal buyValue = buyValueIdx >= 0 ? getBigDecimal(sheet, r, buyValueIdx) : BigDecimal.ZERO;
            BigDecimal sellValue = sellValueIdx >= 0 ? getBigDecimal(sheet, r, sellValueIdx) : BigDecimal.ZERO;
            if (exitDate == null || profit == null) {
                continue; // blank spacer row inside/at end of section
            }

            String tradeTerm = switch (currentSection) {
                case "Equity - Short Term", "Equity - Long Term" -> ZERODHA_SECTION_TERM.get(currentSection);
                default -> { // Non-equity / Mutual Funds: no explicit sub-section, derive from holding period.
                    // Same threshold for every instrument type here, including gold — see CapitalGainsConstants.
                    Long days = holdingIdx >= 0 ? getLong(sheet, r, holdingIdx) : null;
                    yield (days != null && days > CapitalGainsConstants.LONG_TERM_HOLDING_DAYS_THRESHOLD)
                            ? "Long Term" : "Short Term";
                }
            };

            String securityType = switch (currentSection) {
                case "Equity - Short Term", "Equity - Long Term" -> "Equity";
                case "Mutual Funds" -> "Mutual Fund";
                case "Non-equity" -> classifyByName(symbol);
                default -> classifyByName(symbol);
            };

            String quarter = toFinancialQuarter(exitDate);
            results.add(new ParsedRow(
                    symbol.trim(), securityType, tradeTerm, profit, "zerodha",
                    buyValue != null ? buyValue : BigDecimal.ZERO,
                    sellValue != null ? sellValue : BigDecimal.ZERO,
                    quarter));
        }
        return results;
    }

    private boolean isZerodhaSectionHeader(String value) {
        return switch (value) {
            case "Equity - Intraday", "Equity - Short Term", "Equity - Long Term",
                 "Equity - Buyback", "Non-equity", "Mutual Funds", "F&O", "Currency", "Commodity" -> true;
            default -> false;
        };
    }

    // ---------- Groww parser ----------

    private static final Map<String, String> GROWW_SECTION_TERM = Map.of(
            "Short Term trades", "Short Term",
            "Long Term trades", "Long Term"
    );
    private static final Set<String> GROWW_SKIP_SECTIONS = Set.of(
            "Intraday trades", "Buyback trades"
    );

    private List<ParsedRow> parseGroww(Workbook workbook) {
        List<ParsedRow> results = new ArrayList<>();
        Sheet sheet = workbook.getSheet("Sheet1");
        if (sheet == null) return results;

        String currentSection = null;
        boolean inDataBlock = false;

        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            String col0 = getString(sheet, r, 0);
            if (col0 == null || col0.isBlank()) {
                inDataBlock = false;
                continue;
            }
            if (isGrowwSectionHeader(col0)) {
                currentSection = col0;
                inDataBlock = false;
                continue;
            }
            if ("Stock name".equals(col0)) {
                inDataBlock = true;
                continue;
            }
            if (!inDataBlock || currentSection == null || GROWW_SKIP_SECTIONS.contains(currentSection)) {
                continue;
            }

            // Columns: Stock name, ISIN, Quantity, Buy date, Buy price, Buy value,
            //          Sell date, Sell price, Sell value, Realised P&L, Remark
            //
            // Groww already buckets each row into its own "Short Term trades" /
            // "Long Term trades" section using exactly the >365-day rule, for
            // every instrument type including gold — so, unlike Zerodha (which
            // lumps everything non-equity into one undifferentiated section and
            // needs the holding-period math done manually below), no gold
            // override is needed here. The section itself already is the answer.
            String stockName = col0;
            BigDecimal profit = getBigDecimal(sheet, r, 9);
            BigDecimal buyValue = getBigDecimal(sheet, r, 5);
            BigDecimal sellValue = getBigDecimal(sheet, r, 8);
            LocalDate sellDate = parseDate(sheet, r, 6);
            if (sellDate == null || profit == null) continue;

            String tradeTerm = GROWW_SECTION_TERM.get(currentSection);
            String securityType = classifyByName(stockName);

            String quarter = toFinancialQuarter(sellDate);
            results.add(new ParsedRow(
                    stockName.trim(), securityType, tradeTerm, profit, "groww",
                    buyValue != null ? buyValue : BigDecimal.ZERO,
                    sellValue != null ? sellValue : BigDecimal.ZERO,
                    quarter));
        }
        return results;
    }

    private boolean isGrowwSectionHeader(String value) {
        return switch (value) {
            case "Intraday trades", "Short Term trades", "Long Term trades", "Buyback trades" -> true;
            default -> false;
        };
    }

    // ---------- Shared: security type heuristic ----------

    /**
     * Best-effort classification from the instrument name only. Extend these
     * keyword lists as you notice misclassified securities in your output.
     */
    private String classifyByName(String name) {
        // Gold takes priority - must be checked first.
        if (isGoldRelated(name)) {
            return "GOLD";
        }
        String upper = name.toUpperCase(Locale.ROOT);
        if (upper.contains("ETF") || upper.contains("SILVER BEES")
                || upper.contains("BEES")) {
            return "Other";
        }
        if (upper.contains(" MF") || upper.contains("MUTUAL FUND") || upper.contains("AMC")
                || upper.contains("FUND")) {
            return "Mutual Fund";
        }
        return "Equity";
    }

    private String truncateSecurityName(String name) {
        if (name == null || name.isBlank()) return name;

        // Remove hyphens entirely so they are not counted as words
        String cleaned = name.replace("-", "").trim();
        if (cleaned.isBlank()) return cleaned;

        // Split on whitespace and take up to first 4 words
        String[] words = cleaned.split("\\s+");
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < Math.min(4, words.length); i++) {
            if (sb.length() > 0) sb.append(' ');
            sb.append(words[i]);
        }

        String out = sb.toString();
        // If within limit, return as-is
        if (out.length() <= 24) return out;

        // Otherwise truncate and append ellipsis, keeping total length == 24
        int allowed = 24 - 3; // reserve space for "..."
        if (allowed <= 0) return "...";
        String truncated = out.substring(0, Math.min(allowed, out.length()));
        return truncated + "...";
    }

    /**
     * Business rule: gold-related securities are identified by the presence
     * of "GOLD" in their name. Holding-period classification for gold now
     * uses the same {@link CapitalGainsConstants#LONG_TERM_HOLDING_DAYS_THRESHOLD}
     * as every other instrument type — no separate gold threshold.
     */
    private boolean isGoldRelated(String name) {
        String upper = name.toUpperCase(Locale.ROOT);
        return upper.contains("GOLD");
    }

    // ---------- Date / quarter helpers ----------

    private static final List<DateTimeFormatter> DATE_FORMATS = List.of(
            DateTimeFormatter.ofPattern("yyyy-MM-dd"),  // Zerodha
            DateTimeFormatter.ofPattern("dd-MM-yyyy")   // Groww
    );

    /** Indian financial year quarter: Apr-Jun=Q1, Jul-Sep=Q2, Oct-Dec=Q3, Jan-Mar=Q4. */
    private String toFinancialQuarter(LocalDate date) {
        return switch (date.getMonthValue()) {
            case 4, 5, 6 -> "Q1";
            case 7, 8, 9 -> "Q2";
            case 10, 11, 12 -> "Q3";
            default -> "Q4"; // 1, 2, 3
        };
    }

    // ---------- Low-level cell helpers ----------

    private String getString(Sheet sheet, int rowIdx, int colIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return null;
        Cell cell = row.getCell(colIdx);
        if (cell == null) return null;
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? cell.getLocalDateTimeCellValue().toLocalDate().toString()
                    : new DataFormatter().formatCellValue(cell).trim();
            case FORMULA -> new DataFormatter().formatCellValue(cell).trim();
            default -> null;
        };
    }

    private List<String> readRowAsStrings(Sheet sheet, int rowIdx) {
        Row row = sheet.getRow(rowIdx);
        List<String> values = new ArrayList<>();
        if (row == null) return values;
        for (int c = 0; c < row.getLastCellNum(); c++) {
            values.add(Objects.requireNonNullElse(getString(sheet, rowIdx, c), ""));
        }
        return values;
    }

    private LocalDate parseDate(Sheet sheet, int rowIdx, int colIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return null;
        Cell cell = row.getCell(colIdx);
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return cell.getLocalDateTimeCellValue().toLocalDate();
            }
            String raw = getString(sheet, rowIdx, colIdx);
            if (raw == null || raw.isBlank()) return null;
            for (DateTimeFormatter fmt : DATE_FORMATS) {
                try {
                    return LocalDate.parse(raw, fmt);
                } catch (Exception ignored) {
                    // try next format
                }
            }
        } catch (Exception e) {
            return null;
        }
        return null;
    }

    private BigDecimal getBigDecimal(Sheet sheet, int rowIdx, int colIdx) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return null;
        Cell cell = row.getCell(colIdx);
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return BigDecimal.valueOf(cell.getNumericCellValue());
            }
            String raw = getString(sheet, rowIdx, colIdx);
            if (raw == null || raw.isBlank()) return null;
            return new BigDecimal(raw.replace(",", "").trim());
        } catch (Exception e) {
            return null;
        }
    }

    private Long getLong(Sheet sheet, int rowIdx, int colIdx) {
        BigDecimal bd = getBigDecimal(sheet, rowIdx, colIdx);
        return bd == null ? null : bd.longValue();
    }

    private Sheet findSheetStartingWith(Workbook workbook, String prefix) {
        for (Sheet sheet : workbook) {
            if (sheet.getSheetName().startsWith(prefix)) {
                return sheet;
            }
        }
        return null;
    }
}