import { ChangeDetectorRef, Component, OnInit } from '@angular/core';
import { TodoTaskList, TodoTask } from "@microsoft/microsoft-graph-types";
import { Client } from "@microsoft/microsoft-graph-client";

@Component({
  selector: 'app-todo',
  templateUrl: './todo.component.html',
  styleUrls: ['./todo.component.css']
})
export class TodoComponent implements OnInit {

  taskLists: TodoTaskList[] = [];

  constructor(private graphClient: Client, private changeRef: ChangeDetectorRef) { }

  async ngOnInit(): Promise<void> {
    var result = await this.graphClient.api('/me/todo/lists?$filter=displayName eq \'finance\'').get();
    var taskLists = result.value as TodoTaskList[];
    this.taskLists = taskLists;
    await Promise.all(taskLists.map(tl => this.loadList(tl)));
  }

  async loadList(taskList: TodoTaskList) {
    var tasks = await this.graphClient.api(`/me/todo/lists/${taskList.id}/tasks?$filter=status eq 'notStarted'`).get();
    var todoTasks = tasks.value as TodoTask[];
    taskList.tasks = todoTasks;
  }
}
