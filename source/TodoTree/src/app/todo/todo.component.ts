import { Component, OnInit } from '@angular/core';
import { TodoTaskList } from "@microsoft/microsoft-graph-types";
import { Client } from "@microsoft/microsoft-graph-client";

@Component({
  selector: 'app-todo',
  templateUrl: './todo.component.html',
  styleUrls: ['./todo.component.css']
})
export class TodoComponent implements OnInit {

  output = {};
  taskLists: TodoTaskList[] = [];

  constructor(private graphClient: Client) { }

  async ngOnInit(): Promise<void> {


    var result = await this.graphClient.api('/me/todo/lists?$expand=tasks&$top=5').get();

    var taskLists = result.value as TodoTaskList[];
    for (const taskList of taskLists) {
      var result2 = await  this.graphClient.api(`/me/todo/lists/${taskList.id}/tasks?$top=5`).get();
      taskList.tasks = result2.value;
    }

    this.taskLists = taskLists;
}
}
