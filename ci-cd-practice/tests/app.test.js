const request = require('supertest');
const app = require('../src/app');

describe('Task API 集成测试', () => {
  // 每个测试前重置数据
  beforeEach(async () => {
    await request(app).post('/api/reset');
  });

  // ========== 健康检查 ==========
  describe('GET /api/health', () => {
    it('应返回健康状态', async () => {
      const res = await request(app).get('/api/health');
      expect(res.status).toBe(200);
      expect(res.body.status).toBe('ok');
      expect(res.body.timestamp).toBeDefined();
    });
  });

  // ========== 获取任务列表 ==========
  describe('GET /api/tasks', () => {
    it('应返回所有任务', async () => {
      const res = await request(app).get('/api/tasks');
      expect(res.status).toBe(200);
      expect(res.body).toHaveLength(2);
      expect(res.body[0].title).toBe('学习 CI/CD 基础');
    });
  });

  // ========== 获取单个任务 ==========
  describe('GET /api/tasks/:id', () => {
    it('应返回指定任务', async () => {
      const res = await request(app).get('/api/tasks/1');
      expect(res.status).toBe(200);
      expect(res.body.title).toBe('学习 CI/CD 基础');
    });

    it('任务不存在时应返回 404', async () => {
      const res = await request(app).get('/api/tasks/999');
      expect(res.status).toBe(404);
      expect(res.body.error).toBe('任务不存在');
    });
  });

  // ========== 创建任务 ==========
  describe('POST /api/tasks', () => {
    it('应创建新任务', async () => {
      const res = await request(app)
        .post('/api/tasks')
        .send({ title: '部署到生产环境' });
      expect(res.status).toBe(201);
      expect(res.body.id).toBe(3);
      expect(res.body.title).toBe('部署到生产环境');
      expect(res.body.done).toBe(false);
    });

    it('标题为空时应返回 400', async () => {
      const res = await request(app)
        .post('/api/tasks')
        .send({ title: '' });
      expect(res.status).toBe(400);
      expect(res.body.error).toBe('标题不能为空');
    });

    it('缺少标题字段时应返回 400', async () => {
      const res = await request(app)
        .post('/api/tasks')
        .send({});
      expect(res.status).toBe(400);
    });
  });

  // ========== 更新任务 ==========
  describe('PATCH /api/tasks/:id', () => {
    it('应更新任务标题', async () => {
      const res = await request(app)
        .patch('/api/tasks/1')
        .send({ title: '学习 CI/CD 进阶' });
      expect(res.status).toBe(200);
      expect(res.body.title).toBe('学习 CI/CD 进阶');
    });

    it('应更新任务完成状态', async () => {
      const res = await request(app)
        .patch('/api/tasks/1')
        .send({ done: true });
      expect(res.status).toBe(200);
      expect(res.body.done).toBe(true);
    });

    it('任务不存在时应返回 404', async () => {
      const res = await request(app)
        .patch('/api/tasks/999')
        .send({ done: true });
      expect(res.status).toBe(404);
    });
  });

  // ========== 删除任务 ==========
  describe('DELETE /api/tasks/:id', () => {
    it('应删除指定任务', async () => {
      const res = await request(app).delete('/api/tasks/1');
      expect(res.status).toBe(204);

      // 验证已删除
      const check = await request(app).get('/api/tasks/1');
      expect(check.status).toBe(404);
    });

    it('任务不存在时应返回 404', async () => {
      const res = await request(app).delete('/api/tasks/999');
      expect(res.status).toBe(404);
    });
  });
});
