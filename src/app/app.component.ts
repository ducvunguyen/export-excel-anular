import {Component} from '@angular/core';
import { ExcelService } from './excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  data = [
    [1, 2007, 1, 'Volkswagen ', 'Volkswagen Passat', 1267, 10],
    [1, 2007, 1, 'Toyota ', 'Toyota Rav4', 819, 6.5],
    [1, 2007, 1, 'Toyota ', 'Toyota Avensis', 787, 6.2],
    [1, 2007, 1, 'Volkswagen ', 'Volkswagen Golf', 720, 5.7],
    [1, 2007, 1, 'Toyota ', 'Toyota Corolla', 691, 5.4],
    [1, 2007, 1, 'Peugeot ', 'Peugeot 307', 481, 3.8],
    [1, 2008, 1, 'Toyota ', 'Toyota Prius', 217, 2.2],
    [2, 2008, 1, 'Skoda ', 'Skoda Octavia', 216, 2.2],
    [2,2008, 1, 'Peugeot ', 'Peugeot 308', 135, 1.4],
    [2,2008, 2, 'Ford ', 'Ford Mondeo', 624, 5.9],
    [2,2008, 2, 'Volkswagen ', 'Volkswagen Passat', 551, 5.2],
    [2,2008, 2, 'Volkswagen ', 'Volkswagen Golf', 488, 4.6],
    [2,2008, 2, 'Volvo ', 'Volvo V70', 392, 3.7],
    [2,2008, 2, 'Toyota ', 'Toyota Auris', 342, 3.2],
    [2,2008, 2, 'Volkswagen ', 'Volkswagen Tiguan', 340, 3.2],
    [2,2008, 2, 'Toyota ', 'Toyota Avensis', 315, 3],
    [2,2008, 2, 'Nissan ', 'Nissan Qashqai', 272, 2.6],
    [2,2008, 2, 'Nissan ', 'Nissan X-Trail', 271, 2.6],
    [3,2008, 2, 'Mitsubishi ', 'Mitsubishi Outlander', 257, 2.4],
    [3,2008, 2, 'Toyota ', 'Toyota Rav4', 250, 2.4],
    [3,2008, 2, 'Ford ', 'Ford Focus', 235, 2.2],
    [3,2008, 2, 'Skoda ', 'Skoda Octavia', 225, 2.1],
    [4,2008, 2, 'Toyota ', 'Toyota Yaris', 222, 2.1],
    [5,2008, 2, 'Honda ', 'Honda CR-V', 219, 2.1],
    [5,2008, 2, 'Audi ', 'Audi A4', 200, 1.9],
    [5,2008, 2, 'BMW ', 'BMW 3-serie', 184, 1.7],
    [6,2008, 2, 'Toyota ', 'Toyota Prius', 165, 1.6],
    [6,2008, 2, 'Peugeot ', 'Peugeot 207', 144, 1.4]
  ];


  header = ['Year', 'Month', 'Make', 'Model', 'Quantity', 'Pct'];
  title = 'Car Sell Report';

  constructor(private excelService: ExcelService) {

  }

  generateExcel() {
    let test = [];
    this.data.forEach(item => test.push(item[0]));
    this.excelService.generateExcel(this.title, this.header, this.data, 'test_excel', test);
  }

}
