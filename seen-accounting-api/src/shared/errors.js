export class AppError extends Error {
  constructor(message, statusCode = 500, code = 'INTERNAL_ERROR') {
    super(message);
    this.statusCode = statusCode;
    this.code = code;
  }
}

export class NotFoundError extends AppError {
  constructor(entity = 'Resource') {
    super(`${entity} غير موجود`, 404, 'NOT_FOUND');
  }
}

export class ValidationError extends AppError {
  constructor(message = 'بيانات غير صالحة') {
    super(message, 400, 'VALIDATION_ERROR');
  }
}

export class UnauthorizedError extends AppError {
  constructor(message = 'غير مصرح') {
    super(message, 401, 'UNAUTHORIZED');
  }
}

export class ForbiddenError extends AppError {
  constructor(message = 'لا تملك صلاحية') {
    super(message, 403, 'FORBIDDEN');
  }
}
