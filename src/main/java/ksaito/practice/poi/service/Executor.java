package ksaito.practice.poi.service;

import java.io.IOException;
import java.nio.file.Path;
import java.util.concurrent.atomic.AtomicInteger;

import lombok.Builder;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static java.nio.file.Files.newInputStream;
import static java.nio.file.Files.newOutputStream;
import static java.nio.file.Files.readAllLines;
import static java.nio.file.Paths.get;
import static java.util.regex.Pattern.quote;
import static java.util.stream.Collectors.toList;

@Builder
@Slf4j
public class Executor {
  public void run(String[] args) {
    log.info("開始");
    Path template = get("./template.xlsx");
    Path generated = get("./generated.xlsx");
    Path csvData = get("./data.csv");

    log.info("新しいエクセル作成開始");
    try (
      val workbook = new XSSFWorkbook(template.toFile())
    ){
      workbook.write(newOutputStream(generated));
      log.info("新しいエクセル作成終了");
    } catch (IOException e) {
      log.error("ファイル関連エラー", e);
    } catch (InvalidFormatException e) {
      log.error("コンストラクタエラー", e);
    }

    log.info("データ取得・書き込み開始");
    // B14〜
    try (
      val workbook = WorkbookFactory.create(newInputStream(generated))
    ){
      val sheet = workbook.getSheet("template");
      log.info("シート名：" + sheet.getSheetName() + "を取得");
      val rowNumber = new AtomicInteger(2);
      val limit = 10000;
      val recordList = readAllLines(csvData).stream().skip(1).collect(toList());
      log.info("全体で" + recordList.size() + "行");
      val remainder = recordList.size() % limit;
      val divideCount = (recordList.size() - remainder) / limit;
      for (int rowCount = 1; rowCount <= divideCount; rowCount++) {
        recordList.stream()
          .limit(10001)
          .map(record -> record.split(quote(","))
          ).forEach(data -> {
          log.info(rowNumber.get() + "行目");
          val row = sheet.createRow(rowNumber.getAndIncrement());
          for (int cellNumber = 1; cellNumber < 15; cellNumber++) {
            val cell = row.createCell(cellNumber);
            cell.setCellValue(data[cellNumber - 1]);
          }
        });
      }
      workbook.write(newOutputStream(generated));
    } catch (IOException e) {
      log.error("書き込みエラー", e);
    }
    log.info("書き込み終了");
    log.info("終了");
  }
}
