import bcrypt from 'bcryptjs';
import { getDB } from '../../config/database.js';
import { UnauthorizedError, NotFoundError } from '../../shared/errors.js';
import { z } from 'zod';

const loginSchema = z.object({
  email: z.string().email('بريد إلكتروني غير صالح'),
  password: z.string().min(6, 'كلمة المرور يجب أن تكون 6 أحرف على الأقل'),
});

const registerSchema = loginSchema.extend({
  name: z.string().min(1, 'الاسم مطلوب'),
  phone: z.string().optional(),
});

export async function authRoutes(app) {
  const db = () => getDB();

  // ==================== تسجيل الدخول ====================
  app.post('/login', async (request) => {
    const { email, password } = loginSchema.parse(request.body);

    const user = await db().user.findFirst({
      where: { email, isActive: true },
    });
    if (!user || !user.passwordHash) {
      throw new UnauthorizedError('بريد إلكتروني أو كلمة مرور خاطئة');
    }

    const valid = await bcrypt.compare(password, user.passwordHash);
    if (!valid) {
      throw new UnauthorizedError('بريد إلكتروني أو كلمة مرور خاطئة');
    }

    // Update last login
    await db().user.update({
      where: { id: user.id },
      data: { lastLogin: new Date() },
    });

    const token = app.jwt.sign({
      id: user.id,
      name: user.name,
      email: user.email,
      permAdmin: user.permAdmin,
      permReview: user.permReview,
      permTraditionalBot: user.permTraditionalBot,
      permAiBot: user.permAiBot,
      permSheet: user.permSheet,
    });

    return {
      token,
      user: {
        id: user.id,
        name: user.name,
        email: user.email,
        permAdmin: user.permAdmin,
        permReview: user.permReview,
      },
    };
  });

  // ==================== تسجيل مستخدم جديد ====================
  app.post('/register', async (request, reply) => {
    const { name, email, password, phone } = registerSchema.parse(request.body);

    const passwordHash = await bcrypt.hash(password, 12);
    const user = await db().user.create({
      data: { name, email, passwordHash, phone },
    });

    const token = app.jwt.sign({
      id: user.id,
      name: user.name,
      email: user.email,
      permAdmin: user.permAdmin,
      permReview: user.permReview,
      permTraditionalBot: user.permTraditionalBot,
      permAiBot: user.permAiBot,
      permSheet: user.permSheet,
    });

    return reply.status(201).send({
      token,
      user: { id: user.id, name: user.name, email: user.email },
    });
  });

  // ==================== الملف الشخصي ====================
  app.get('/me', { preHandler: [app.authenticate] }, async (request) => {
    const user = await db().user.findUnique({
      where: { id: request.user.id },
      select: {
        id: true,
        name: true,
        email: true,
        phone: true,
        telegramUsername: true,
        permAdmin: true,
        permReview: true,
        permTraditionalBot: true,
        permAiBot: true,
        permSheet: true,
        lastLogin: true,
        createdAt: true,
      },
    });
    if (!user) throw new NotFoundError('المستخدم');
    return user;
  });

  // ==================== تحديث كلمة المرور ====================
  app.put('/password', { preHandler: [app.authenticate] }, async (request) => {
    const { currentPassword, newPassword } = z
      .object({
        currentPassword: z.string(),
        newPassword: z.string().min(6),
      })
      .parse(request.body);

    const user = await db().user.findUnique({ where: { id: request.user.id } });
    const valid = await bcrypt.compare(currentPassword, user.passwordHash);
    if (!valid) {
      throw new UnauthorizedError('كلمة المرور الحالية خاطئة');
    }

    const passwordHash = await bcrypt.hash(newPassword, 12);
    await db().user.update({
      where: { id: request.user.id },
      data: { passwordHash },
    });

    return { message: 'تم تغيير كلمة المرور بنجاح' };
  });
}
