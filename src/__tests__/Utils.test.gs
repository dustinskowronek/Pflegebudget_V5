require.extensions['.gs'] = require.extensions['.js'];
require('../Utils.gs');

describe('Utils.slug', () => {
  test('konvertiert Namen zu Slug', () => {
    expect(Utils.slug('Hans Müller')).toBe('hans-müller');
  });
});
