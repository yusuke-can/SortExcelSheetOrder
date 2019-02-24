package com.commandline.SortExcelSheetOrder;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.PathMatcher;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import javax.validation.Validation;
import javax.validation.ValidationException;
import javax.validation.Validator;
import javax.validation.ValidatorFactory;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.dom4j.DocumentException;
import org.dom4j.Node;
import org.dom4j.io.SAXReader;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.util.StringUtils;
import org.yaml.snakeyaml.Yaml;

import com.commandline.SortExcelSheetOrder.SortExcelSheetOrderApplication.ConfigYmlDto.ClasspathFile;

import lombok.Data;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

@Slf4j
@SpringBootApplication
public class SortExcelSheetOrderApplication {

    @Data
    public static class ConfigYmlDto {

        /**
         * .classpathファイルの相対パス:classpathFileRelativePath が設定された時にチェックするグループ
         * @author aplag
         *
         */
        interface ClasspathFile {}

        /**
         * ビルド対象packageのみフラグ
         * <pre>
         * src/test/java配下のビルド対象のpackageみ対象とする場合、true
         * src/test/java配下の全packageを対象とする場合、false
         *
         * .classpathに記載されている src/test/javaの設定を参照する
         * </pre>
         */
        @NotNull
        private Boolean onlyBuildTargetPackage;

        /**
         * シート名順序ファイル相対パス
         */
        @NotBlank
        private String sheetNameOrderFileRelativePath;

        /**
         * デフォルト処理対象ディレクトリ
         */
        @NotBlank
        private String defaultTargetDirectory;

        /**
         * ファイル名globパターン
         */
        @NotBlank
        private String globFileNamePattern;

        /**
         * プロジェクトルートからの.classpathの相対パス
         */
        private String classpathFileRelativePath;

        /**
         * デフォルト処理対象パッケージ
         */
        @NotBlank(groups=ClasspathFile.class)
        private String defaultTargetPackage;

        /**
         * 除外packageリスト
         */
        private List<String> excludePackageList;
    }

    @SneakyThrows
    public static void main(String[] args) {
        SpringApplication.run(SortExcelSheetOrderApplication.class, args);

        log.info("\n" + IntStream.range(0, args.length)
                .mapToObj(i -> String.format("arg[%d]: %s", i, args[i]))
                .collect(Collectors.joining("\n")));

        /*
         *  バッチ引数
         */
        Path configYmlPath = Paths.get(args[0]);
        if (!Files.exists(configYmlPath)) {
            throw new FileNotFoundException("configYmlPath: " + configYmlPath.toString());
        }

        /*
         * YAMLファイルからパラメタ取得
         */
        Yaml yaml = new Yaml();
        ConfigYmlDto configYmlDto = yaml.loadAs(Files.newInputStream(configYmlPath), ConfigYmlDto.class);
        log.info(configYmlDto.toString());
        validateConfigYmlDto(configYmlDto);

        // プロジェクトルートのディレクトリパス
        Path projectRootPath = Paths.get(System.getProperty("user.dir"));
        // デフォルトの処理対象ディレクトリ
        Path defaultTargetDirectoryPath = Paths.get(projectRootPath.toString(), configYmlDto.getDefaultTargetDirectory());
        if (!Files.exists(defaultTargetDirectoryPath)) {
            throw new FileNotFoundException("defaultTargetDirectoryPath: " + defaultTargetDirectoryPath.toString());
        }

        // 処理対象ディレクトリのパスリスト
        List<Path> targetDirectoryPathList = getTargetDirectoryPathList(projectRootPath, defaultTargetDirectoryPath, configYmlDto);
        log.info("\n" + IntStream.range(0, targetDirectoryPathList.size())
                .mapToObj(i -> String.format("[処理対象ディレクトリパス] targetDirectoryPathList[%d]: %s", i, targetDirectoryPathList.get(i)))
                .collect(Collectors.joining("\n")));

        targetDirectoryPathList.stream().map(Path::toString).forEach(log::info);

        Path sheetNameOrderFilePath = Paths.get(projectRootPath.toString(), configYmlDto.getSheetNameOrderFileRelativePath());
        if (!Files.exists(sheetNameOrderFilePath)) {
            throw new FileNotFoundException("sheetNameOrderFilePath: " + sheetNameOrderFilePath.toString());
        }

        List<String> sheetNameOrderList = Files.readAllLines(sheetNameOrderFilePath);
        validateDuplicateSheetName(sheetNameOrderFilePath, sheetNameOrderList);

        Map<String, Integer> sheetNameOrderMap = IntStream.range(0, sheetNameOrderList.size())
              .boxed()
              .collect(Collectors.toMap(sheetNameOrderList::get ,Integer::valueOf));

        PathMatcher filePathMatcher = FileSystems.getDefault().getPathMatcher("glob:" + configYmlDto.getGlobFileNamePattern());
        targetDirectoryPathList.parallelStream().forEach(targetPath -> {
            try {
                Files.walk(targetPath)
                .parallel()
                .filter(filePathMatcher::matches)
                // Excel, wordファイルを開いている時に生成される一時ファイルを除外
                .filter(path -> !path.getFileName().toString().startsWith("~$"))
                .forEach(path -> sortExcelSheetOrder(path, sheetNameOrderMap));
            } catch (IOException e) {
                e.printStackTrace();
            }
        });
    }

    /**
     * コンフィグファイル(YAML)から取得したパラメタの入力チェック.
     * @param configYmlDto コンフィグYAML DTO
     */
    private static void validateConfigYmlDto(ConfigYmlDto configYmlDto) {
        ValidatorFactory factory = Validation.buildDefaultValidatorFactory();
        Validator validator = factory.getValidator();

        List<String> validationErrorMessageList = validator.validate(configYmlDto).stream()
                .map(cv -> cv.toString()).collect(Collectors.toList());

        /*
         * .classpathの相対パス: classpathFileRelativePath が設定されている場合、
         * .classpathファイル用のパラメタの入力チェックを行う
         */
        if (!StringUtils.isEmpty(configYmlDto.getClasspathFileRelativePath())) {
            validationErrorMessageList.addAll(validator.validate(configYmlDto,ClasspathFile.class).stream()
                    .map(cv -> cv.toString()).collect(Collectors.toList()));
        }

        if (validationErrorMessageList.size() > 0) {
            throw new ValidationException("\n" + validationErrorMessageList.stream().collect(Collectors.joining("\n")));
        }
    }

    /**
     * Excelファイル検索対象のフォルダパスリスト取得.
     *
     * @param classpathFilePath .classpathファイルのパス
     * @param projectRootPath プロジェクトのルートディレクトリパス
     * @param configYmlDto コンフィグファイル（YAML）
     * @return 処理対象のフォルダパスリスト
     * @throws DocumentException
     * @throws IOException
     */
    private static List<Path> getTargetDirectoryPathList(Path projectRootPath, Path defaultTargetDirectoryPath,
            ConfigYmlDto configYmlDto) throws DocumentException, IOException {
        if(!configYmlDto.getOnlyBuildTargetPackage() ) {
            return Files.list(defaultTargetDirectoryPath).collect(Collectors.toList());
        } else if(StringUtils.isEmpty(configYmlDto.getClasspathFileRelativePath())) {
            return Files.list(defaultTargetDirectoryPath).collect(Collectors.toList());
        }

        Path classpathFilePath = Paths.get(projectRootPath.toString(), configYmlDto.getClasspathFileRelativePath());
        if (!Files.exists(classpathFilePath)) {
            return Files.list(defaultTargetDirectoryPath).collect(Collectors.toList());
        }

        SAXReader reader = new SAXReader();

        // .classpathの要素名: classpathentryのうち、path属性の値が引数と一致する要素を取得
        Node srcTestJavaNode = reader.read(classpathFilePath.toString()).getRootElement()
                .selectSingleNode("//classpathentry[@path='" + configYmlDto.getDefaultTargetPackage() + "']");

        // .classpathののinclude属性の値を取得
        String includeAttributeValue = srcTestJavaNode.valueOf("@including");
        if (StringUtils.isEmpty(includeAttributeValue)) {
            return Files.list(defaultTargetDirectoryPath).collect(Collectors.toList());
        }

        // include属性の値はclasspathが複数ある場合はパイプ("|")区切りとなっているため、パイプ("|")で分割して絶対パスに変換
        return Arrays.stream(includeAttributeValue.split("\\|"))
                // 処理対象外のクラスパスは除外
                .filter(classpath -> Optional.ofNullable(configYmlDto.getExcludePackageList())
                        .map(list -> !list.contains(classpath)).orElse(true))
                .map(classpath -> Paths.get(projectRootPath.toString(), configYmlDto.getDefaultTargetPackage(), classpath))
                .filter(Files::exists)
                .filter(Files::isDirectory)
                .collect(Collectors.toList());

    }

    /**
     * シート名順リストに重複したシート名が含まれている場合、例外をスロー
     * @param sheetNameOrderFilePath シート名順ファイルパス
     * @param sheetNameOrderList シート名順リスト
     */
    private static void validateDuplicateSheetName(Path sheetNameOrderFilePath, List<String> sheetNameOrderList) {
        Set<String> sheetNameSet = new HashSet<>(sheetNameOrderList);
        if(sheetNameOrderList.size() == sheetNameSet.size()) {
            return;
        }

        Set<String> sheetNameUniqueSet = new HashSet<>();
        String duplicateSheetNamesSplittedByComma = sheetNameOrderList.stream()
                .filter(sheetName -> !sheetNameUniqueSet.add(sheetName))
                .distinct()
                .collect(Collectors.joining(",","[","]"));

        throw new IllegalArgumentException(String.format("[シート名順ファイル] %s のシート名が重複しています。 重複シート名: %s"
                , sheetNameOrderFilePath.toString(), duplicateSheetNamesSplittedByComma));
    }

    /**
     * Excelブックのシートをシート名順リストの順に左→右に並び替える. シート名順リストにないシートは右側に移動する
     * @param excelFilePath Excelファイルパス
     * @param sheetNameOrderMap シート名順Map(キー：シート名, 値: シートの並び順（左→→右へ1からの連番）)
     */
    @SneakyThrows
    private static void sortExcelSheetOrder(Path excelFilePath, Map<String, Integer> sheetNameOrderMap) {

        try (Workbook book = WorkbookFactory.create(new FileInputStream(excelFilePath.toString()))) {
            if (book.getNumberOfSheets() == 1) {
                log.debug(String.format("file path: %s : シート数が1シートのみのため、ソート不要",
                        excelFilePath.toString()));
                return;
            }

            List<String> sheetNameList = IntStream.range(0, book.getNumberOfSheets())
                    .mapToObj(book::getSheetAt)
                    .map(Sheet::getSheetName)
                    .collect(Collectors.toList());

            TreeMap<Integer, String> orderdSheetNameMap = new TreeMap<>();
            List<String> unknownSheetNameList = new ArrayList<>();
            sheetNameList.stream().forEach(sheetName -> {
                if (sheetNameOrderMap.containsKey(sheetName)) {
                    orderdSheetNameMap.put(sheetNameOrderMap.get(sheetName), sheetName);
                } else {
                    unknownSheetNameList.add(sheetName);
                }
            });
            if (orderdSheetNameMap.isEmpty()) {
                log.debug(String.format("file path: %s : ソート対象のシートが存在しないため、ソート不要\n"
                        + " - [%s] ",
                        excelFilePath.toString(),
                        sheetNameList.stream().collect(Collectors.joining(","))));
                return;
            }

            List<String> orderdSheetNameList = orderdSheetNameMap.values().stream().collect(Collectors.toList());
            orderdSheetNameList.addAll(unknownSheetNameList);
            IntStream.range(0, orderdSheetNameList.size())
                    .forEach(i -> book.setSheetOrder(orderdSheetNameList.get(i), i));
            if (sheetNameList.equals(orderdSheetNameList)) {
                log.debug(String.format("file path: %s : シートの並び順はソート済のため、ソート処理不要\n"
                        + " - [%s] ",
                        excelFilePath.toString(),
                        sheetNameList.stream().collect(Collectors.joining(","))));
                return;
            }

            try (FileOutputStream out = new FileOutputStream(excelFilePath.toString())) {
                book.write(out);
            }
            log.info(String.format("\n"
                    + "file path: %s.\n"
                    + " - before: [%s] \n"
                    + " - after : [%s]",
                    excelFilePath.toString(),
                    sheetNameList.stream().collect(Collectors.joining(",")),
                    orderdSheetNameList.stream().collect(Collectors.joining(","))));
        }
    }

}
