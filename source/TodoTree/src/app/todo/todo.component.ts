import { Component, OnInit } from '@angular/core';
import { TodoTaskList } from "@microsoft/microsoft-graph-types";
import { Client } from "@microsoft/microsoft-graph-client";

@Component({
  selector: 'app-todo',
  templateUrl: './todo.component.html',
  styleUrls: ['./todo.component.css']
})
export class TodoComponent implements OnInit {

  loaded = 0;
  taskLists: TodoTaskList[] | null = null;

  constructor(private graphClient: Client) { }

  async ngOnInit(): Promise<void> {
    var result = await this.graphClient.api('/me/todo/lists?$expand=tasks').get();
    var taskLists = result.value as TodoTaskList[];
    this.taskLists = taskLists;
    await Promise.all(taskLists.map(tl => this.loadList(tl)));
  }

  async loadList(taskList: TodoTaskList) {
    var tasks = await this.graphClient.api(`/me/todo/lists/${taskList.id}/tasks?$top=5`).get();
    taskList.tasks = tasks.value;
    this.loaded++;
  }
}
