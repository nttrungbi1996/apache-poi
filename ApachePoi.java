package apache;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.Units;

public class ApachePoi {

	// 元のExcelファイルのパスに適切に変更します。
	final static String PATH = "C:\\Users\\MGT社員\\Desktop\\aws\\テンプレートファイル.xlsx";
	// 修正内容を保存する新しいExcelファイルのパスに適切に変更します。
	final static String PATH_TARGET = "C:\\Users\\MGT社員\\Desktop\\aws\\sample.xlsx";

	// main　メッソドがプログラムの実行を開始します。
	public static void main(String[] args) {

		// お客様が選択した項目をList<String> valuesに格納します
		List<String> values = new ArrayList<>();
		values.add("(1) 売買");
		values.add("(2) 交換");
		values.add("(3) 代理");
		values.add("(4) 媒介");
		values.add("(1) 契約の締結");
		values.add("(2) 契約の申込みの受理");

		// values (List<String>) を渡して、apachePoiメソッドを呼び出して、Apache POIの操作を実行します。
		apachePoi(values);

	}

	// このメソッドは、Excelファイルに対してApache POIの操作を行います。
	private static void apachePoi(List<String> values) {
		try ( // FileInputStreamを使用して、元のExcelファイル（PATHで指定されたパス）を読み込みます。
				FileInputStream fileIn = new FileInputStream(PATH);
				// WorkbookFactory.create(fileIn)を使用して、読み込んだExcelファイルからワークブックを作成します。
				Workbook workbook = WorkbookFactory.create(fileIn);) {

			// ワークブックから指定したシート名（この場合は"テンプレート"という名前のシート）を取得します。
			Sheet sheet = workbook.getSheet("テンプレート");

			// シート内のすべての行とセルをループ処理
			sheet.forEach(row -> {
				row.forEach(cell -> {

					// 現在のセルの値を取得
					Object originalValue = getCellValue(cell);

					//nullではない場合
					if (originalValue != null) {

						// セルの値がvaluesリストに存在するか確認
						if (values.contains(originalValue)) {

							// 円形の画像をセルに追加するためのDrawingオブジェクトを作成
							Drawing<?> drawing = sheet.createDrawingPatriarch();

							// 円形の画像データ（PNG形式）を作成し、ワークブックに追加し、インデックスを取得
							int pictureIndex = workbook.addPicture(getCircleImageData(workbook),
									Workbook.PICTURE_TYPE_PNG);

							// セル内での画像の位置とサイズを定義
							int dx1 = -Units.EMU_PER_PIXEL * 3;
							int dy1 = Units.EMU_PER_PIXEL * 8;
							int dy2 = Units.EMU_PER_PIXEL * 8;

							// 図形の配置調整
							ClientAnchor anchor = drawing.createAnchor(dx1, dy1, 0, dy2,
									cell.getColumnIndex(), cell.getRowIndex(),
									cell.getColumnIndex(), cell.getRowIndex());

							// セルに円形の画像を挿入
							Picture picture = drawing.createPicture(anchor, pictureIndex);
							picture.resize();
						}

					}
				});
			});

			// 修正した内容を出力ファイルに書き込む
			try (FileOutputStream fileOut = new FileOutputStream(new File(PATH_TARGET))) {
				workbook.write(fileOut);
			} catch (Exception e) {
				e.printStackTrace();
			}
			System.out.println("終了！");

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 円形画像のサイズを取得
	private static int getCircleSize() {
		return 30; // 円形画像のサイズ（30x30ピクセル）を返します
	}

	// 黒い境界線の円形画像を作成し、その画像データ（PNG形式）を返す
	private static byte[] getCircleImageData(Workbook workbook) {

		int size = getCircleSize(); // 円形画像のサイズを取得
		BufferedImage image = new BufferedImage(size, size, BufferedImage.TYPE_INT_ARGB);
		Graphics2D g2d = image.createGraphics();
		g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

		// 透明な円を描画
		g2d.setColor(new Color(0, 0, 0, 0)); // 透明な色で設定
		g2d.fillOval(0, 0, size, size);

		// 黒い境界線の円を描画
		g2d.setColor(Color.BLACK);
		int strokeWidth = 1; // 輪の太さ
		g2d.setStroke(new BasicStroke(strokeWidth));
		g2d.drawOval(strokeWidth / 2, strokeWidth / 2, size - strokeWidth, size - strokeWidth);

		g2d.dispose();

		ByteArrayOutputStream baos = new ByteArrayOutputStream();

		try {
			ImageIO.write(image, "png", baos);
		} catch (IOException e) {
			e.printStackTrace();
		}

		return baos.toByteArray();
	}

	// セルの値を取得
	private static Object getCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getRichStringCellValue().getString();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				return cell.getNumericCellValue();
			}
		case BOOLEAN:
			return cell.getBooleanCellValue();
		case FORMULA:
			return cell.getCellFormula();
		default:
			return null;
		}
	}
}
