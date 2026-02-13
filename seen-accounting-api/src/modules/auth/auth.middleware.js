import { ForbiddenError } from '../../shared/errors.js';

// Check specific permission
export function requirePermission(permission) {
  return async (request, reply) => {
    await request.jwtVerify();
    const user = request.user;

    const permMap = {
      admin: 'permAdmin',
      review: 'permReview',
      traditional_bot: 'permTraditionalBot',
      ai_bot: 'permAiBot',
      sheet: 'permSheet',
    };

    const field = permMap[permission];
    if (!field || !user[field]) {
      throw new ForbiddenError(`لا تملك صلاحية: ${permission}`);
    }
  };
}

// Admin only
export function requireAdmin() {
  return requirePermission('admin');
}
