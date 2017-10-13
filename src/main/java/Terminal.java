import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;


public class Terminal {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        TreeMap<String, Integer> lectures = new TreeMap<>();
        try {
            students = SheetUtil.getAllStudents();
            lectures = SheetUtil.getAllLectures();
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
        System.out.println("Select one of the following lectures (you have to type it!):");
        for (String s : lectures.keySet()) {
            System.out.println(s);
        }
        String selectedLessonName;

        do {
            BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
            selectedLessonName = reader.readLine();
            if (lectures.containsKey(selectedLessonName)) {
                System.out.println("You have selected " + selectedLessonName);
                reader.close();
                break;
            } else {
                System.out.println(selectedLessonName + " not found, try again!");
            }
        } while (true);
        int[] lessonInfo = SheetUtil.getLectureInfoByColumnId(lectures.get(selectedLessonName));
        int lessonColumnId = lessonInfo[0];
        int lessonJudgeId = lessonInfo[1];
        int lessonMaxPoints = lessonInfo[2];

        //region Adding my personal cookies
        cookies.put("_ym_uid", "");
        cookies.put("__RequestVerificationToken", "");
        cookies.put("ASP.NET_SessionId", "");
        cookies.put("_ym_isad", "");
        cookies.put(".AspNet.SoftUniJudgeCookie", "");
        cookies.put("_ga", "");
        cookies.put("_gid", "");
        //endregion

        Document root = getDocument("/Contests/Practice/Results/Simple/" + lessonJudgeId);
        String linkToNextPage = getLinkToNextPage(root);
        while (linkToNextPage != null) {
            Elements tableRows = root.select("tr");
            for (Element tableRow : tableRows) {
                String user = tableRow.select("td a").text();
                Integer userRowId = students.get(user);
                if (userRowId != null) {
                    String[] parts = tableRow.select("td:last-child").text().split("\\s+");
                    int points = Integer.parseInt(parts[0]);
                    double assessment = 6.0 / lessonMaxPoints * points;
                    System.out.println("Found student " + user + " with points " + points + " which is " + assessment);
                    SheetUtil.writeResultToCell(userRowId, lessonColumnId, assessment);
                    System.out.println("Successfully added assessment for student " + user + " at worksheet.");
                    students.replace(user, -1);
                }

            }
            if (toStop()) break;

            root = getDocument(linkToNextPage);
            linkToNextPage = getLinkToNextPage(root);
        }
        System.out.println("Ready, view your file!");
    }

    private static Map<String, Integer> students = new HashMap<>();

    private static Map<String, String> cookies = new HashMap<>();

    private static int page = 1;

    private static Document getDocument(String url) {
        Document doc = null;
        try {
            String baseUrl = "https://judge.softuni.bg";
            System.out.println("Requesting data for page " + page);
            doc = Jsoup.connect(baseUrl + url).cookies(cookies).get();
            System.out.println("Fetching data for page " + page++ + " done.");
        } catch (IOException e) {
            System.out.println("Connection timed out. Retrying...");
            // Retry
            getDocument(url);
        }
        return doc;
    }

    private static String getLinkToNextPage(Document doc) {
        String linkToNextPage = doc.select("li.pull-right a").first().attr("href");
        if (linkToNextPage.equals("#")) return null;
        return linkToNextPage;
    }

    private static boolean toStop() {
        for (Integer val : students.values()) {
            if (val >= 0) return false;
        }
        return true;
    }
}
