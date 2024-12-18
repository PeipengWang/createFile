package textdemo;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class FileGeneratorController {
    public static void main(String[] args) {
        List<String> lists = new ArrayList<>();
        lists.add("aaa");
        lists.add("bbb");
        generateFile("E:/code/createFile/test", lists);
    }

    /**
     * REST API 生成文件
     * @param path 文件生成的目录路径
     * @param params 文件内容参数（可以多个）
     * @return 生成结果信息
     */
    public static String generateFile(String path, List<String> params) {

        // 确保路径末尾无多余斜杠
        File baseDir = new File(path.endsWith(File.separator) ? path : path + File.separator);

        // 创建以时间戳命名的文件夹
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMddHHmmss"));
        File timestampDir = new File(baseDir, timestamp);

        if (!timestampDir.exists()) {
            boolean created = timestampDir.mkdirs();
            if (!created) {
                return "Error: Could not create directory: " + timestampDir.getAbsolutePath();
            }
        }

        // 创建文件
        File txtFile = new File(timestampDir, "data.txt");
        try (FileWriter writer = new FileWriter(txtFile)) {
            for (String param : params) {
                writer.write(param + System.lineSeparator());
            }
        } catch (IOException e) {
            return "Error: Could not write to file: " + txtFile.getAbsolutePath();
        }

        return "File created successfully at: " + txtFile.getAbsolutePath();
    }
}
