package org.example;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.stage.Stage;


/**
 * JavaFX entry point.
 */
public class FxApp extends Application {

    @Override
    public void start(Stage stage) {
        MainView view = new MainView();

        Scene scene = new Scene(view.getRoot(), 940, 860);

        view.applyStyles(scene);

        stage.setTitle("조리지시서 변환기");
        stage.setScene(scene);
        stage.setMinWidth(780);
        stage.setMinHeight(600);
        stage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
