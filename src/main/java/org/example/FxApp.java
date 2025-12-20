package org.example;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.net.URL;

/**
 * JavaFX entry point.
 */
public class FxApp extends Application {

    @Override
    public void start(Stage stage) {
        MainView view = new MainView();

        Scene scene = new Scene(view.getRoot(), 860, 560);

        URL css = getClass().getResource("/styles.css");
        if (css != null) {
            scene.getStylesheets().add(css.toExternalForm());
        } else {
            System.err.println("styles.css not found");
        }

        view.applyStyles(scene);

        stage.setTitle("조리지시서 변환기");
        stage.setScene(scene);
        stage.setMinWidth(760);
        stage.setMinHeight(520);
        stage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
