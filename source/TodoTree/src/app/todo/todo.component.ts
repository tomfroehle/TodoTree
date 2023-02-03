import { ChangeDetectorRef, Component, OnInit } from '@angular/core';
import { TodoTaskList, TodoTask } from "@microsoft/microsoft-graph-types";
import { Client } from "@microsoft/microsoft-graph-client";
import {graphviz} from "d3-graphviz";
@Component({
  selector: 'app-todo',
  templateUrl: './todo.component.html',
  styleUrls: ['./todo.component.css']
})
export class TodoComponent implements OnInit {


  constructor(private graphClient: Client) { }

  async ngOnInit(): Promise<void> {
    var result = await this.graphClient.api('/me/todo/lists?$filter=displayName eq \'finance\'').get();
    var taskLists = result.value as TodoTaskList[];
    await this.loadList(taskLists[0]);

    const nodes = taskLists[0].tasks?.map((t,i)=> `node${i} [label="${t.title}"];`).join(' ');
    const edges = taskLists[0].tasks?.map((t,i)=> `node${i}`).join(' ');

    graphviz('#graph').renderDot(`digraph {
      layout=neato
      ${nodes}
      ${edges}
    }`);
  }

  async loadList(taskList: TodoTaskList) {
    var tasks = await this.graphClient.api(`/me/todo/lists/${taskList.id}/tasks?$filter=status eq 'notStarted'`).get();
    var todoTasks = tasks.value as TodoTask[];
    taskList.tasks = todoTasks;
  }
}
