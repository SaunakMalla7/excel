import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  ngOnInit(): void {
    console.log(this.groupList[1].studentList[4].phone);
  }
  title = 'excel';
  fileName = 'ExcelSheet.xlsx';

  groupList: any[] = [
    {
      'group': 'L1C1',
      'studentList': [
        {
          id: 1,
          name: "Granville",
          phone: '6515611',
          email: 'Granville@gmail.com',
        },
        {
          id: 2,
          name: 'Audrey A. Fredericks',
          phone: '859-948-6255',
          email: 'AudreyAFredericks@jourrapide.com',
        },
        {
          id: 3,
          name: 'Javier V. Jackson',
          phone: '618-659-3734',
          email: 'JavierVJackson@armyspy.com',
        },
        {
          id: 4,
          name: 'Gloria V. Tabron',
          phone: '478-812-6082',
          email: 'GloriaVTabron@rhyta.com',
        },
        {
          id: 5,
          name: 'Arlene R. Cooper',
          phone: '312-533-5687',
          email: 'ArleneRCooper@armyspy.com',
        },
        {
          id: 6,
          name: 'Earl J. Roman',
          phone: '801-754-8265',
          email: 'EarlJRoman@teleworm.us',
        },
        {
          id: 7,
          name: 'Mary R. Baltazar',
          phone: '319-483-5677',
          email: 'MaryRBaltazar@teleworm.us',
        },
        {
          id: 8,
          name: 'Larry R. Williams',
          phone: '727-343-5653',
          email: 'LarryRWilliams@armyspy.com',
        },
        {
          id: 9,
          name: 'Carolyn E. Watters',
          phone: '904-579-2974',
          email: 'CarolynEWatters@armyspy.com',
        },
        {
          id: 10,
          name: 'Ralph T. Jones',
          phone: '678-714-2360',
          email: 'RalphTJones@rhyta.com',
        },
      ]
    },
    {
      group: 'L1C2',
      studentList: [
        {
          id: 1,
          name: 'Fiore Zetticci',
          phone: '0350 7525146',
          email: 'FioreZetticci@teleworm.us',
        },
        {
          id: 2,
          name: 'Fausto Pinto',
          phone: '0341 2702500',
          email: 'FaustoPinto@armyspy.com',
        },
        {
          id: 3,
          name: 'Alvisa Palermo',
          phone: '0355 0973605',
          email: 'AlvisaPalermo@armyspy.com',
        },
        {
          id: 4,
          name: 'Gloria V. Tabron',
          phone: '0385 0669713',
          email: 'GloriaVTabron@rhyta.com',
        },
        {
          id: 5,
          name: 'Mareta Padovesi',
          phone: '0393 3743568',
          email: 'MaretaPadovesi@jourrapide.com',
        },
        {
          id: 6,
          name: 'Filippa Onio',
          phone: '0356 3560203',
          email: 'FilippaOnio@dayrep.com',
        },
        {
          id: 7,
          name: 'Mary R. Baltazar',
          phone: '319-483-5677',
          email: 'MaryRBaltazar@teleworm.us',
        },
        {
          id: 8,
          name: 'Larry R. Williams',
          phone: '727-343-5653',
          email: 'LarryRWilliams@armyspy.com',
        },
        {
          id: 9,
          name: 'Carolyn E. Watters',
          phone: '904-579-2974',
          email: 'CarolynEWatters@armyspy.com',
        },
        {
          id: 10,
          name: 'Ralph T. Jones',
          phone: '678-714-2360',
          email: 'RalphTJones@rhyta.com',
        },
      ]
    },
    {
      group: 'L1C3',
      studentList: [
        {
          id: 1,
          name: "Reece Kay",
          phone: "078 8646 0139",
          email: "ReeceKay@rhyta.com"
        },
        {
          id: 2,
          name: "Ellie Palmer",
          phone: "078 6097 0693",
          email: "ElliePalmer@einrot.com"
        },
        {
          id: 3,
          name: "Mia Parker",
          phone: "077 5067 2982",
          email: "MiaParker@cuvox.de"
        },
        {
          id: 4,
          name: "Louis Vincent",
          phone: "079 8748 6044",
          email: "LouisVincent@teleworm.us"
        },
        {
          id: 5,
          name: "Ryan Parsons",
          phone: "078 4877 0816",
          email: "RyanParsons@armyspy.com"
        },
        {
          id: 6,
          name: 'Filippa Onio',
          phone: '0356 3560203',
          email: 'FilippaOnio@dayrep.com',
        },
        {
          id: 7,
          name: 'Mary R. Baltazar',
          phone: '319-483-5677',
          email: 'MaryRBaltazar@teleworm.us',
        },
        {
          id: 8,
          name: 'Larry R. Williams',
          phone: '727-343-5653',
          email: 'LarryRWilliams@armyspy.com',
        },
        {
          id: 9,
          name: 'Carolyn E. Watters',
          phone: '904-579-2974',
          email: 'CarolynEWatters@armyspy.com',
        },
        {
          id: 10,
          name: 'Ralph T. Jones',
          phone: '678-714-2360',
          email: 'RalphTJones@rhyta.com',
        },
      ]
    }
  ]

  exportexcel(): void {

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    this.groupList.forEach((obj: any) => {
      // sort by name
      obj.studentList = obj.studentList.sort((a: any, b: any) => (a.name > b.name) ? 1 : -1)
      const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(obj.studentList);

      XLSX.utils.book_append_sheet(wb, ws, obj.group);
    });

    /* save to file */
    XLSX.writeFile(wb, this.fileName);


  }
}
