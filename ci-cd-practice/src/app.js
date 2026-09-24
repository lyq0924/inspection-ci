const express = require('express');

const app = express();
app.use(express.json());

// 内存数据存储
let tasks = [
  { id: 1, title: '学习 CI/CD 基础', done: false },
  { id: 2, title: '配置 GitHub Actions', done: false },
];
let nextId = 3;

// 健康检查
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// 获取所有任务
app.get('/api/tasks', (req, res) => {
  res.json(tasks);
});

// 获取单个任务
app.get('/api/tasks/:id', (req, res) => {
  const id = parseInt(req.params.id, 10);
  const task = tasks.find(t => t.id === id);
  if (!task) {
    return res.status(404).json({ error: '任务不存在' });
  }
  res.json(task);
});

// 创建任务
app.post('/api/tasks', (req, res) => {
  const { title } = req.body;
  if (!title || typeof title !== 'string' || title.trim() === '') {
    return res.status(400).json({ error: '标题不能为空' });
  }
  const task = { id: nextId++, title: title.trim(), done: false };
  tasks.push(task);
  res.status(201).json(task);
});

// 更新任务状态
app.patch('/api/tasks/:id', (req, res) => {
  const id = parseInt(req.params.id, 10);
  const task = tasks.find(t => t.id === id);
  if (!task) {
    return res.status(404).json({ error: '任务不存在' });
  }
  if (req.body.title !== undefined) {
    if (typeof req.body.title !== 'string' || req.body.title.trim() === '') {
      return res.status(400).json({ error: '标题不能为空' });
    }
    task.title = req.body.title.trim();
  }
  if (req.body.done !== undefined) {
    task.done = Boolean(req.body.done);
  }
  res.json(task);
});

// 删除任务
app.delete('/api/tasks/:id', (req, res) => {
  const id = parseInt(req.params.id, 10);
  const index = tasks.findIndex(t => t.id === id);
  if (index === -1) {
    return res.status(404).json({ error: '任务不存在' });
  }
  tasks.splice(index, 1);
  res.status(204).send();
});

// 重置数据（用于测试）
app.post('/api/reset', (req, res) => {
  tasks = [
    { id: 1, title: '学习 CI/CD 基础', done: false },
    { id: 2, title: '配置 GitHub Actions', done: false },
  ];
  nextId = 3;
  res.json({ message: '数据已重置' });
});

module.exports = app;
