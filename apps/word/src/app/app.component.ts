import { Component } from '@angular/core';

@Component({
  selector: 'word-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  constructor() {
    console.log("App component initialized");

  }
  title = 'Welcome to the Word plugin';
}
