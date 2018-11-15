package com.github.houbb.iexcel.cache;

/**
 * @author binbin.hou
 * @date 2018/11/15 18:11
 */
public interface ICache<K, V> {

    /**
     * 从缓存池中查找值
     *
     * @param key 键
     * @return 值
     */
    V get(K key);

    /**
     * 放入缓存
     *
     * @param key   键
     * @param value 值
     * @return 值
     */
    V put(K key, V value);

    /**
     * 移除缓存
     *
     * @param key 键
     * @return 移除的值
     */
    V remove(K key);

    /**
     * 清空缓存池
     */
    void clear();

}
